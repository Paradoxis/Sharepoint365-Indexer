package main

import (
	"crypto/tls"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path"
	"strconv"
	"strings"
	"sync"
	"time"
)

//
// Utility types
//

type CrawlerHttpTransport struct {
	base      http.RoundTripper
	cookies   map[string]string
	userAgent string
	proxyURL  *url.URL
	insecure  bool
}

func (t *CrawlerHttpTransport) RoundTrip(req *http.Request) (*http.Response, error) {
	req2 := req.Clone(req.Context())
	req2.Header.Set("User-Agent", t.userAgent)

	for name, value := range t.cookies {
		req2.AddCookie(&http.Cookie{Name: name, Value: value})
	}

	transport := &http.Transport{}

	if t.proxyURL != nil {
		transport.Proxy = http.ProxyURL(t.proxyURL)
	}

	if t.insecure {
		transport.TLSClientConfig = &tls.Config{InsecureSkipVerify: true}
	}

	return transport.RoundTrip(req2)
}

//
// JSON structures
//

type SharePointContextWebInformation struct {
	D struct {
		GetContextWebInformation struct {
			Metadata struct {
				Type string `json:"type"`
			} `json:"__metadata"`
			FormDigestTimeoutSeconds int    `json:"FormDigestTimeoutSeconds"`
			FormDigestValue          string `json:"FormDigestValue"`
			LibraryVersion           string `json:"LibraryVersion"`
			SiteFullUrl              string `json:"SiteFullUrl"`
			SupportedSchemaVersions  struct {
				Metadata struct {
					Type string `json:"type"`
				} `json:"__metadata"`
				Results []string `json:"results"`
			} `json:"SupportedSchemaVersions"`
			WebFullUrl string `json:"WebFullUrl"`
		} `json:"GetContextWebInformation"`
	} `json:"d"`
}

type TokenResponse struct {
	ODataContext  string  `json:"@odata.context"`
	ODataType     string  `json:"@odata.type"`
	ODataID       string  `json:"@odata.id"`
	ODataEditLink string  `json:"@odata.editLink"`
	AccessToken   string  `json:"access_token"`
	ExpiresOn     string  `json:"expires_on"`
	IDToken       *string `json:"id_token"`
	Resource      string  `json:"resource"`
	Scope         string  `json:"scope"`
	TokenType     string  `json:"token_type"`
}

type SharePointSiteDetails struct {
	ODataContext  string `json:"@odata.context"`
	IsInDraftMode bool   `json:"IsInDraftMode"`
	IsVivaBackend bool   `json:"IsVivaBackend"`
	SiteId        string `json:"SiteId"`
	WebId         string `json:"WebId"`
	LogoUrl       string `json:"LogoUrl"`
	Title         string `json:"Title"`
	Url           string `json:"Url"`
}

type SharePointSite struct {
	Id         int              `json:"Id"`
	ParentId   int              `json:"ParentId"`
	Title      string           `json:"Title"`
	Url        string           `json:"Url"`
	IsExternal bool             `json:"IsExternal"`
	Chidlren   []SharePointSite `json:"Children"`
}

type PrettySharePointFod struct {
	Id          string      `json:"Id"`
	Permissions Permissions `json:"Permissions"`
	Type        string      `json:"Type"`
	Path        string      `json:"Path"`
	FullPath    string      `json:"FullPath"`
	Size        int64       `json:"Size"`
	Site        string      `json:"Site"`
}

func (fod *PrettySharePointFod) String() string {
	return fmt.Sprintf("%s (%d bytes)", fod.FullPath, fod.Size)
}

func (fod *PrettySharePointFod) Format(format string) string {

	switch format {
	case "text":
		return fod.String()
	case "jsonl":
		data, err := json.Marshal(fod)
		if err != nil {
			log.Printf("Failed to marshal PrettySharePointFod: %v", err)
			return ""
		}

		return string(data)
	}

	log.Fatalln("Unsupported format:", format)
	return ""
}

type SharePointFod struct {
	Id      string   `json:"UniqueId"`
	Mask    string   `json:"PermMask"`
	Type    string   `json:"FSObjType"`
	Path    string   `json:"FileRef"`
	Size    string   `json:"SMTotalSize"`
	Site    *string  `json:"Site,omitempty"`
	SiteURL *url.URL `json:"SiteUrl,omitempty"`
}

type SharePointStatResponse struct {
	D *struct {
		Metadata struct {
			Type string `json:"type"`
		} `json:"__metadata"`
	} `json:"d,omitempty"`

	Error *struct {
		Code    string `json:"code"`
		Message struct {
			Value string `json:"value"`
		} `json:"message"`
	} `json:"error,omitempty"`
}

func (fod *SharePointFod) FullPath() string {
	if fod.SiteURL == nil {
		panic("SiteURL is nil, cannot construct full path")
	}

	siteURL := *fod.SiteURL
	site := siteURL.Scheme + "://" + siteURL.Host
	return site + fod.Path
}

func (fod *SharePointFod) String() string {
	return fmt.Sprintf("%s (%s bytes)", fod.FullPath(), fod.Size)
}

func (fod *SharePointFod) Format(format string) string {
	size, err := strconv.ParseInt(fod.Size, 10, 64)
	if err != nil {
		log.Printf("Failed to parse size for SharePointFod: %v", err)
		size = 0 // Default to 0 if parsing fails
	}

	mask, err := strconv.ParseUint(strings.TrimPrefix(fod.Mask, "0x"), 16, 64)
	if err != nil {
		log.Printf("Failed to parse mask for SharePointFod: %v", err)
		mask = 0 // Default to 0 if parsing fails
	}

	pretty := PrettySharePointFod{
		Id:          strings.Trim(fod.Id, "{}"),
		Type:        fodTypeToString(fod.Type),
		Size:        size,
		Path:        fod.Path,
		FullPath:    fod.FullPath(),
		Permissions: ParsePermissionMask(mask),
		Site:        "",
	}

	if fod.Site != nil {
		pretty.Site = *fod.Site
	}

	return pretty.Format(format)
}

const FILE_TYPE_FOLDER = "1"
const FILE_TYPE_FILE = "0"

func fodTypeToString(fodType string) string {
	switch fodType {
	case FILE_TYPE_FOLDER:
		return "directory"
	case FILE_TYPE_FILE:
		return "file"
	default:
		return "unknown"
	}
}

type Permissions struct {
	Mask          uint64   `json:"Mask"`
	SymbolicNames []string `json:"SymbolicNames"`
}

// permissionsMap maps symbolic names to their bitmask values
// source: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wssfo3/1f5e3322-920f-431c-bbc4-7f65c477e698
var permissionsMap = map[string]uint64{

	// List and document permissions
	"ViewListItems":             0x0000000000000001,
	"AddListItems":              0x0000000000000002,
	"EditListItems":             0x0000000000000004,
	"DeleteListItems":           0x0000000000000008,
	"ApproveItems":              0x0000000000000010,
	"OpenItems":                 0x0000000000000020,
	"ViewVersions":              0x0000000000000040,
	"DeleteVersions":            0x0000000000000080,
	"CancelCheckout":            0x0000000000000100,
	"ManagePersonalViews":       0x0000000000000200,
	"ManageLists":               0x0000000000000800,
	"ViewFormPages":             0x0000000000001000,
	"AnonymousSearchAccessList": 0x0000000000002000,

	// Web level permissions
	"Open":                          0x0000000000010000,
	"ViewPages":                     0x0000000000020000,
	"AddAndCustomizePages":          0x0000000000040000,
	"ApplyThemeAndBorder":           0x0000000000080000,
	"ApplyStyleSheets":              0x0000000000100000,
	"ViewUsageData":                 0x0000000000200000,
	"CreateSSCSite":                 0x0000000000400000,
	"ManageSubwebs":                 0x0000000000800000,
	"CreateGroups":                  0x0000000001000000,
	"ManagePermissions":             0x0000000002000000,
	"BrowseDirectories":             0x0000000004000000,
	"BrowseUserInfo":                0x0000000008000000,
	"AddDelPrivateWebParts":         0x0000000010000000,
	"UpdatePersonalWebParts":        0x0000000020000000,
	"ManageWeb":                     0x0000000040000000,
	"AnonymousSearchAccessWebLists": 0x0000000080000000,

	// Extended permissions
	"UseClientIntegration": 0x0000001000000000,
	"UseRemoteAPIs":        0x0000002000000000,
	"ManageAlerts":         0x0000004000000000,
	"CreateAlerts":         0x0000008000000000,
	"EditMyUserInfo":       0x0000010000000000,

	// Special permissions
	"EnumeratePermissions": 0x4000000000000000,
}

func ParsePermissionMask(mask uint64) Permissions {
	var symbolicNames []string

	for name, value := range permissionsMap {
		if mask&value != 0 {
			symbolicNames = append(symbolicNames, name)
		}
	}

	return Permissions{
		Mask:          mask,
		SymbolicNames: symbolicNames,
	}
}

type SharePointFodResponse struct {
	Row      []SharePointFod `json:"Row"`
	FirstRow int             `json:"FirstRow"`
	LastRow  int             `json:"LastRow"`
}

//
// Crawlerf
//

type SharePointCookies struct {
	RtFa    string `json:"rtFa"`
	FedAuth string `json:"FedAuth"`
}

type SharePointConfig struct {
	ProxyURL  *url.URL
	BaseURL   string
	Cookies   SharePointCookies
	UserAgent string
	Insecure  bool
}

type SharePointCrawler struct {
	Config SharePointConfig
	client *http.Client
}

func NewSharePointCrawler(config SharePointConfig) *SharePointCrawler {
	client := &http.Client{}

	client.Transport = &CrawlerHttpTransport{
		base:      http.DefaultTransport,
		userAgent: config.UserAgent,
		cookies: map[string]string{
			"rtFa":    config.Cookies.RtFa,
			"FedAuth": config.Cookies.FedAuth,
		},
		proxyURL: config.ProxyURL,
		insecure: config.Insecure,
	}

	if client.Transport == nil || client.Transport == http.DefaultTransport {
		client.Transport = &CrawlerHttpTransport{
			base:      http.DefaultTransport,
			userAgent: config.UserAgent,
		}
	}

	return &SharePointCrawler{
		Config: config,
		client: client,
	}
}

func (sc *SharePointCrawler) GetConfig() SharePointConfig {
	return sc.Config
}

func (sc *SharePointCrawler) GetFormDigestValue() (string, error) {
	var contextInfo SharePointContextWebInformation

	req, err := http.NewRequest("POST", sc.Config.BaseURL+"/_api/contextinfo", nil)
	if err != nil {
		return "", err
	}

	req.Header.Set("Accept", "application/json;odata=verbose")
	req.Header.Set("Content-Type", "application/json;odata=none")

	resp, err := sc.client.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	err = json.NewDecoder(resp.Body).Decode(&contextInfo)
	if err != nil {
		return "", err
	}

	if contextInfo.D.GetContextWebInformation.FormDigestValue == "" {
		return "", fmt.Errorf("form digest value is empty")
	}

	return contextInfo.D.GetContextWebInformation.FormDigestValue, nil
}

func (sc *SharePointCrawler) GetAccessToken(resource string) (TokenResponse, error) {
	var tokenResponse TokenResponse

	digest, err := sc.GetFormDigestValue()
	if err != nil {
		return tokenResponse, fmt.Errorf("failed to get form digest value: %v", err)
	}

	req, err := http.NewRequest("POST", sc.Config.BaseURL+"/_api/SP.OAuth.Token/Acquire", nil)
	if err != nil {
		return tokenResponse, err
	}

	req.Header.Set("X-RequestDigest", digest)
	req.Header.Set("Accept", "application/json;odata.metadata=minimal")
	req.Header.Set("Content-Type", "application/json; charset=UTF-8")
	req.Header.Set("Odata-Version", "4.0")

	payload := map[string]string{"resource": resource}
	payloadBytes, err := json.Marshal(payload)
	if err != nil {
		return tokenResponse, fmt.Errorf("failed to marshal payload: %v", err)
	}
	req.Body = io.NopCloser(strings.NewReader(string(payloadBytes)))

	resp, err := sc.client.Do(req)
	if err != nil {
		return tokenResponse, fmt.Errorf("failed to execute request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return tokenResponse, fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	err = json.NewDecoder(resp.Body).Decode(&tokenResponse)
	if err != nil {
		return tokenResponse, fmt.Errorf("failed to decode response: %v", err)
	}

	if tokenResponse.AccessToken == "" {
		return tokenResponse, fmt.Errorf("access token is empty")
	}

	return tokenResponse, nil
}

type Office365ResultSet struct {
	Results []Office365ResultItem `json:"Results"`
	Total   int64                 `json:"Total"`
}

type Office365EntitySet struct {
	ResultSets []Office365ResultSet `json:"ResultSets"`
}

type Office365SearchResults struct {
	EntitySets []Office365EntitySet `json:"EntitySets"`
}

type Office365ResultItem struct {
	Id     string                `json:"Id"`
	Type   string                `json:"Type"`
	Source Office365SearchResult `json:"Source"`
}

type Office365SearchResult struct {
	IsDocument bool   `json:"isDocument"`
	FileType   string `json:"FileType"`
	Size       int64  `json:"Size"`
	Path       string `json:"Path"`
	Write      string `json:"Write"`
	SiteId     string `json:"SiteId"`
	SiteName   string `json:"SiteName"`
	UniqueId   string `json:"UniqueId"`
}

type Office365Cursor struct {
	Total  int64                   `json:"total"`
	Offset int64                   `json:"offset"`
	Size   int64                   `json:"size"`
	Files  []Office365SearchResult `json:"files"`
}

func (sc *SharePointCrawler) SearchOffice365(query string, offset int64, size int64, sort map[string]string) (Office365Cursor, error) {
	var cursor Office365Cursor
	var searchResponse Office365SearchResults

	accessToken, err := sc.GetAccessToken("https://outlook.office365.com/search")
	if err != nil {
		return cursor, fmt.Errorf("failed to retrieve access token")
	}

	req, err := http.NewRequest("POST", "https://outlook.office365.com/searchservice/api/v2/query", nil)
	if err != nil {
		return cursor, fmt.Errorf("failed to create request: %v", err)
	}

	req.Header.Set("Authorization", "Bearer "+accessToken.AccessToken)
	req.Header.Set("Accept", "application/json")
	req.Header.Set("Content-Type", "application/json")

	searchRequest := NewSearchRequest(query, offset, size, sort)
	payload, err := json.Marshal(searchRequest)
	if err != nil {
		return cursor, fmt.Errorf("failed to marshal search request: %v", err)
	}
	req.Body = io.NopCloser(strings.NewReader(string(payload)))

	resp, err := sc.client.Do(req)
	if err != nil {
		return cursor, fmt.Errorf("failed to execute search request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return cursor, fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	err = json.NewDecoder(resp.Body).Decode(&searchResponse)
	if err != nil {
		return cursor, fmt.Errorf("failed to decode search response: %v", err)
	}

	if len(searchResponse.EntitySets) == 0 || len(searchResponse.EntitySets[0].ResultSets) == 0 {
		return cursor, fmt.Errorf("no results found in search response")
	}
	resultSet := searchResponse.EntitySets[0].ResultSets[0]

	var result []Office365SearchResult = make([]Office365SearchResult, 0)

	for _, item := range resultSet.Results {
		result = append(result, item.Source)
	}

	cursor.Files = result
	cursor.Total = resultSet.Total
	cursor.Offset = offset
	cursor.Size = 100
	return cursor, nil
}

func NewSearchRequest(query string, from int64, size int64, sort map[string]string) interface{} {
	sortFields := []map[string]interface{}{}
	for field, direction := range sort {
		sortFields = append(sortFields, map[string]interface{}{
			"Field":         field,
			"SortDirection": direction,
		})
	}

	return map[string]interface{}{
		"Scenario": map[string]interface{}{
			"Name": "SPHomeWeb",
			"Dimensions": []map[string]interface{}{
				{
					"DimensionName":  "QueryType",
					"DimensionValue": "Files",
				},
			},
		},
		"Cvid": "7f8970e0-b383-4387-86e7-3190b379db79",
		"EntityRequests": []map[string]interface{}{
			{
				"EntityType": "File",
				"ContentSources": []string{
					"SharePoint",
					"OneDriveBusiness",
				},
				"From": from,
				"Size": size,
				"Fields": []string{
					"isDocument",
					"FileType",
				},
				"Query": map[string]interface{}{
					"QueryString":        "*",
					"DisplayQueryString": "*",
					"QueryTemplate":      query,
					// "QueryTemplate":      "({searchterms}) AND IsContainer:true AND isDocument:false AND ContentClass:STS_Site",
					// "QueryTemplate":      "({searchterms}) isDocument:true AND NOT (FileType:onetoc2 OR FileType:asp OR FileType:aspx OR FileType:htm OR FileType:html OR FileType:mhtml) ",
				},
				"RefiningQueries": []map[string]interface{}{
					{
						"RefinerString": "LastModifiedTime:(range(min, 2100-01-01T00:00:00.000Z))",
					},
				},
				"Sort": sortFields,
			},
		},
	}
}

func (sc *SharePointCrawler) GetSites(attempts int) ([]string, error) {
	var sites map[string]bool = make(map[string]bool)
	var query string = "({searchterms}) AND IsContainer:true AND isDocument:false AND ContentClass:STS_Site"
	var size int64 = 100
	var sort map[string]string = map[string]string{
		"PersonalScore": "Desc",
	}

	for i := 0; i < attempts; i++ {
		var offset int64 = 0

		items, err := sc.SearchOffice365(query, offset, 100, sort)
		if err != nil {
			if strings.Contains(err.Error(), "unexpected status code: 500") {
				continue // sharepoint be like that sometimes
			}

			log.Println("Failed to list sites:", err)
			break
		}

		for offset < items.Total {
			offset += items.Size

			for _, file := range items.Files {
				parsedUrl, err := url.Parse(file.Path)
				if err != nil {
					log.Println("Failed to parse file path:", file.Path, err)
					continue
				}

				if !strings.HasPrefix(parsedUrl.Path, "/sites/") {
					continue
				}

				sites[file.Path] = true
			}

			if offset >= items.Total {
				break
			}

			items, err = sc.SearchOffice365(query, offset, size, sort)
			if err != nil {
				log.Println("Failed to list sites:", err)
				break
			}

			log.Println("Attempt", i+1, "found", len(sites), "total sites")
			time.Sleep(2 * time.Second)
		}
	}

	var siteList []string
	for site := range sites {
		siteList = append(siteList, site)
	}

	return siteList, nil
}

func (sc *SharePointCrawler) ListDocumentsRecursive(site string, directory string, depth int, output chan SharePointFod) error {
	defer func() {
		if depth == 0 {
			close(output)
		}
	}()

	documents, err := sc.ListDocuments(site, directory)
	if err != nil {
		return fmt.Errorf("failed to site root: %v", err)
	}

	for _, doc := range documents {
		output <- doc

		if doc.Type == FILE_TYPE_FOLDER {
			err = sc.ListDocumentsRecursive(site, doc.Path, depth+1, output)
			if err != nil {
				log.Printf("Failed to list documents in folder %s: %v", doc.Path, err)
			}
		}
	}

	return nil
}

func (sc *SharePointCrawler) ListDocuments(site string, directory string) ([]SharePointFod, error) {
	siteURL, err := url.Parse(site)
	if err != nil {
		return nil, fmt.Errorf("failed to parse site URL: %v", err)
	}

	endpoint := fmt.Sprintf("%s/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%s", site, url.QueryEscape(fmt.Sprintf("'%s/Shared Documents'", siteURL.Path)))
	if directory != "" {
		endpoint += "&RootFolder=" + url.QueryEscape(directory)
	}

	req, err := http.NewRequest("POST", endpoint, nil)
	if err != nil {
		return nil, err
	}

	req.Header.Set("Accept", "application/json;odata=none")
	req.Header.Set("Content-Type", "application/json;odata=verbose")
	req.Header.Set("Odata-Version", "4.0")

	resp, err := sc.client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("failed to execute request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	var fodResponse SharePointFodResponse
	err = json.NewDecoder(resp.Body).Decode(&fodResponse)
	if err != nil {
		return nil, fmt.Errorf("failed to decode response: %v", err)
	}

	for i := range fodResponse.Row {
		fodResponse.Row[i].Site = &site
		fodResponse.Row[i].SiteURL = siteURL
	}

	return fodResponse.Row, nil
}

func (sc *SharePointCrawler) GetDocument(site string, path string) (io.ReadCloser, error) {
	endpoint := fmt.Sprintf("%s/_api/web/GetFileByServerRelativePath(DecodedUrl=@a1)/OpenBinaryStream?@a1=%s", site, url.QueryEscape(fmt.Sprintf("'%s'", path)))

	req, err := http.NewRequest("GET", endpoint, nil)
	if err != nil {
		return nil, err
	}

	req.Header.Set("Accept", "application/json;odata=verbose")
	req.Header.Set("Content-Type", "application/json;odata=verbose")

	resp, err := sc.client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("failed to execute request: %v", err)
	}

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	return resp.Body, nil
}

func (sc *SharePointCrawler) IsFile(site string, path string) (bool, error) {
	endpoint := fmt.Sprintf("%s/_api/web/GetFileByServerRelativePath(DecodedUrl=@a1)?$select=__metadata&@a1=%s", site, url.QueryEscape(fmt.Sprintf("'%s'", path)))

	req, err := http.NewRequest("GET", endpoint, nil)
	if err != nil {
		return false, err
	}

	req.Header.Set("Accept", "application/json;odata=verbose")
	req.Header.Set("Content-Type", "application/json;odata=verbose")

	resp, err := sc.client.Do(req)
	if err != nil {
		return false, fmt.Errorf("failed to execute request: %v", err)
	}

	if !(resp.StatusCode == http.StatusOK || resp.StatusCode == http.StatusNotFound) {
		return false, fmt.Errorf("unexpected status code: %d", resp.StatusCode)
	}

	var statResponse SharePointStatResponse
	err = json.NewDecoder(resp.Body).Decode(&statResponse)
	if err != nil {
		return false, fmt.Errorf("failed to decode response: %v", err)
	}

	if statResponse.D != nil {
		if statResponse.D.Metadata.Type == "SP.File" {
			return true, nil
		}

		if statResponse.D.Metadata.Type == "SP.Folder" {
			return false, nil
		}
	}

	if statResponse.Error != nil {
		if strings.Contains(statResponse.Error.Message.Value, "does not exist") {
			return false, nil
		}

		return false, fmt.Errorf("got unexpected error message whilst performing stat on: %s: %s", site+path, statResponse.Error.Message.Value)
	}

	return false, fmt.Errorf("unexpected response structure: %v", statResponse)
}

//
// Main entrypoint
//

func main() {

	baseUrl := flag.String("url", "", "Base URL of the SharePoint site (e.g., https://example.sharepoint.com)")
	cookies := flag.String("cookies", "", "Cookie string containing rtFa and FedAuth cookies")
	output := flag.String("output", "", "Output file for the results")
	format := flag.String("format", "text", "Console output format (none, text or jsonl)")
	outputFormat := flag.String("fformat", "text", "File output format (text or jsonl)")
	site := flag.String("site", "", "Specific SharePoint site to crawl (optional)")
	download := flag.String("download", "", "Download a given file or directory from SharePoint (best effort)")
	shared := flag.Bool("shared", false, "Attempt to look for files shared with the user directly")
	list := flag.Bool("list", false, "List all SharePoint sites instead of crawling")
	quiet := flag.Bool("quiet", false, "Suppress other console output")
	proxy := flag.String("proxy", "", "Proxy URL for the HTTP client (e.g., http://127.0.0.1:8080, socks5://127.0.0.1:1080)")
	insecure := flag.Bool("insecure", false, "Skip TLS verification for the HTTP client (not recommended)")
	threads := flag.Int("threads", 4, "Number of concurrent threads to use for crawling")
	discoveryAttempts := flag.Int("discovery-attempts", 2, "Number of attempts to try to discover SharePoint sites (default: 2). The API returns sites randomly and doesn't support indexing past 500 results, so this is a workaround to get more sites.")
	userAgent := flag.String("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/421.24 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0", "User-Agent string for the HTTP client")

	flag.Parse()

	if *cookies == "" {
		flag.Usage()
		return
	}

	if *threads <= 0 {
		log.Fatalln("Number of threads must be greater than 0")
	}

	if *format != "none" && *format != "text" && *format != "jsonl" {
		log.Fatalln("Invalid format specified. Supported formats are 'none', 'text' and 'jsonl'")
	}

	if *outputFormat != "text" && *outputFormat != "jsonl" {
		log.Fatalln("Invalid output format specified. Supported formats are 'text' and 'jsonl'")
	}

	if *quiet {
		log.SetOutput(io.Discard)
	}

	var outputFile *os.File

	if *output != "" {
		file, err := os.Create(*output)
		if err != nil {
			log.Fatalln("Failed to create output file:", err)
		}

		outputFile = file
		defer file.Close()
	}

	if *output == "" && *format == "none" {
		log.Fatalln("Output file must be specified when format is 'none', what's the point of indexing without output?")
	}

	if *site != "" && *list {
		log.Fatalln("Cannot specify both -site and -list flags at the same time")
	}

	if *download != "" && *list {
		log.Fatalln("Cannot specify both -download and -list flags at the same time")
	}

	var proxyURL *url.URL
	if *proxy != "" {
		proxy, err := url.Parse(*proxy)
		if err != nil {
			log.Fatalln("Failed to parse proxy URL:", err)
		}

		proxyURL = proxy
		log.Println("Using proxy:", proxyURL)
	}

	parsedCookies := parseCookies(*cookies)
	sharepointCookies := SharePointCookies{
		RtFa:    parsedCookies["rtFa"],
		FedAuth: parsedCookies["FedAuth"],
	}

	sharepointConfig := SharePointConfig{
		Insecure:  *insecure,
		ProxyURL:  proxyURL,
		BaseURL:   *baseUrl,
		UserAgent: *userAgent,
		Cookies:   sharepointCookies,
	}

	crawler := NewSharePointCrawler(sharepointConfig)
	log.Println("SharePoint Crawler initialized with base URL:", crawler.Config.BaseURL)

	if *download != "" {
		log.Println("Starting download of:", *download)

		downloadURL, err := url.Parse(*download)
		if err != nil {
			log.Fatalln("Failed to parse site URL:", *download)
		}

		if strings.HasPrefix(downloadURL.Path, "/personal/") {
			log.Println("Downloading personal document:", *download)

			outputPath := path.Join(downloadURL.Host, downloadURL.Path)
			log.Println("File will be written to:", outputPath)

			err := os.MkdirAll(path.Dir(outputPath), 0755)
			if err != nil {
				log.Fatalln("Failed to create output directory:", err)
			}

			req, err := http.NewRequest("GET", *download, nil)
			if err != nil {
				log.Fatalf("Failed to create request for personal document: %v", err)
			}

			resp, err := crawler.client.Do(req)
			if err != nil {
				log.Fatalf("Failed to download personal document: %v", err)
			}
			defer resp.Body.Close()

			if resp.StatusCode != http.StatusOK {
				log.Fatalf("Unexpected status code: %d", resp.StatusCode)
			}

			file, err := os.Create(outputPath)
			if err != nil {
				log.Fatalln("Failed to create output file:", err)
			}
			defer file.Close()

			_, err = io.Copy(file, resp.Body)
			if err != nil {
				log.Fatalln("Failed to save downloaded file:", err)
			}

			log.Println("Successfully downloaded personal document to:", outputPath)
			return
		}

		if !strings.HasPrefix(downloadURL.Path, "/sites/") {
			log.Fatalln("Download path must start with /sites/")
		}

		siteBase := strings.SplitN(downloadURL.Path, "/Shared Documents/", 2)
		site := downloadURL.Scheme + "://" + downloadURL.Host + siteBase[0]

		var wg sync.WaitGroup
		results := make(chan SharePointFod)
		semaphore := make(chan bool, *threads)

		isFile, err := crawler.IsFile(site, downloadURL.Path)
		if err != nil {
			log.Fatalln("Failed to check if path is a file:", err)
		}

		if isFile {
			log.Println("Downloading file:", downloadURL.Path)
			go func() {
				defer close(results)

				// kind of a hack but I'm too lazy to refactor the code
				results <- SharePointFod{
					Id:      "",
					Mask:    "",
					Type:    FILE_TYPE_FILE,
					Path:    downloadURL.Path,
					Site:    &site,
					SiteURL: downloadURL,
					Size:    "0",
				}
			}()
		} else {
			go func() {
				err := crawler.ListDocumentsRecursive(site, downloadURL.Path, 0, results)
				if err != nil {
					log.Printf("Error listing documents recursively: %v\n", err)
				} else {
					log.Printf("Finished listing documents for site %s\n", site)
				}
			}()
		}

		for fod := range results {
			wg.Add(1)

			go func(fod SharePointFod) {
				semaphore <- true
				defer wg.Done()
				defer func() { <-semaphore }()

				outputPath := path.Join(downloadURL.Host, fod.Path)

				err := os.MkdirAll(path.Dir(outputPath), 0755)
				if err != nil {
					log.Fatalln("Failed to create output directory:", err)
				}

				if fod.Type == FILE_TYPE_FOLDER {
					err := os.MkdirAll(outputPath, 0755)
					if err != nil {
						log.Fatalln("Failed to create output directory:", err)
					}
					return
				}

				// check if the file already exists
				if _, err := os.Stat(outputPath); err == nil {
					log.Println("File already exists, skipping download:", outputPath)
					return
				}

				// yolo no dir traversal checking
				file, err := os.Create(outputPath)
				if err != nil {
					log.Fatalln("Failed to create output file:", err)
				}

				defer file.Close()

				log.Println("Got successful HTTP response for:", fod.FullPath())
				body, err := crawler.GetDocument(site, fod.Path)
				if err != nil {
					log.Fatalln("Failed to download document:", err)
				}
				defer body.Close()

				log.Println("Saving downloaded file to:", file.Name())
				_, err = io.Copy(file, body)
				if err != nil {
					log.Fatalln("Failed to save downloaded file:", err)
				}
			}(fod)
		}

		wg.Wait()
		log.Println("Channel closed, all documents downloaded")
		return
	}

	if *baseUrl != "" && *site != "" {
		log.Println("Using specified base URL:", *baseUrl)
	}

	if *baseUrl == "" && *site != "" {
		parsedSite, err := url.Parse(*site)
		if err != nil {
			log.Fatalln("Failed to parse site URL:", *site)
		}

		if !strings.HasPrefix(parsedSite.Path, "/sites/") {
			log.Fatalln("Site path must start with /sites/")
		}

		*baseUrl = parsedSite.Scheme + "://" + parsedSite.Host
		log.Println("Inferring base URL from site:", *baseUrl)
	}

	if *baseUrl == "" && *site == "" {
		log.Fatalln("Base URL must be specified with -url flag or -site flag")
	}

	if *shared {
		log.Println("Attempting to look for files shared with the user..")
		var query string = "({searchterms}) AND ContentClass:STS_ListItem_MySiteDocumentLibrary AND isDocument:true"
		var size int64 = 1000000
		var sort map[string]string = map[string]string{
			"Created": "Asc",
		}

		result, err := crawler.SearchOffice365(query, 0, size, sort)
		if err != nil {
			log.Fatalln("Failed to search for shared files:", err)
		}

		deduplicated := make(map[string]bool)

		for _, file := range result.Files {
			if ok, exists := deduplicated[file.UniqueId]; exists && ok {
				continue // skip already processed files
			}

			deduplicated[file.UniqueId] = true

			fod := PrettySharePointFod{
				Id:          strings.Trim(file.UniqueId, "{}"),
				Permissions: ParsePermissionMask(0), // No permissions available for personal files
				Type:        "file",
				FullPath:    file.Path,
				Path:        strings.Replace(file.Path, file.SiteName, "", 1),
				Site:        file.SiteName,
				Size:        file.Size,
			}

			if *format != "none" {
				fmt.Println(fod.Format(*format))
			}

			if outputFile != nil {
				outputFile.WriteString(fod.Format(*outputFormat) + "\n")
			}
		}

		return
	}

	var crawlSites []string

	if *site == "" {
		log.Println("Enumerating sharepoint sites..")
		sites, err := crawler.GetSites(*discoveryAttempts)
		if err != nil {
			log.Fatalln("Failed to retrieve sites:", err)
		}
		log.Println("Found", len(sites), "sites")
		crawlSites = sites
	} else {
		log.Println("Using specified site:", *site)
		crawlSites = []string{*site}
	}

	if *list {
		for _, site := range crawlSites {
			fmt.Println(site)
		}
		return
	}

	var wg sync.WaitGroup
	results := make(chan SharePointFod)
	semaphore := make(chan bool, *threads)

	for _, site := range crawlSites {
		wg.Add(1)

		go func(site string) {
			semaphore <- true
			defer wg.Done()
			defer func() { <-semaphore }()

			log.Println("Indexing sharepoint site:", site)
			intermediate := make(chan SharePointFod)

			go func() {
				err := crawler.ListDocumentsRecursive(site, "", 0, intermediate)
				if err != nil {
					log.Printf("Error listing documents recursively: %v\n", err)
				} else {
					log.Printf("Finished listing documents for site %s\n", site)
				}
			}()

			for doc := range intermediate {
				results <- doc
			}
		}(site)
	}

	go func() {
		defer close(results)
		wg.Wait()

		log.Println("All sites indexed, closing results channel")
	}()

	log.Println("Waiting for results...")

	for fod := range results {
		if *format != "none" {
			fmt.Println(fod.Format(*format))
		}

		if outputFile != nil {
			outputFile.WriteString(fod.Format(*outputFormat) + "\n")
		}
	}

	log.Println("Crawling completed successfully")
}

func parseCookies(cookies string) map[string]string {
	cookieMap := make(map[string]string)
	for _, cookie := range strings.Split(cookies, ";") {
		parts := strings.SplitN(strings.TrimSpace(cookie), "=", 2)
		if len(parts) == 2 {
			cookieMap[parts[0]] = parts[1]
		}
	}
	return cookieMap
}
