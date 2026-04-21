// ============================================================================
// REST_API_Pagination.pq
// Power Query (M) function: consume a paginated REST API inside Excel/Power BI.
//
// WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
// ---------------------------------------
// Excel's "From Web" connector reads ONE URL. It does not follow "next page"
// links, does not handle bearer-token auth (nicely), and chokes on rate
// limits. This M function does all three. You drop it into the Advanced
// Editor of a new query, set the three parameters up top, and get a flat
// table of every record across every page.
//
// USE CASE (software business)
//  - Pull every opportunity out of HubSpot / Salesforce / Pipedrive
//  - Pull every ticket out of Zendesk / Freshdesk / ServiceNow
//  - Pull every order out of Stripe / QuickBooks / Xero
//  - Pull every build out of GitHub Actions / Jenkins
//
// USAGE
// -----
// Open Excel -> Data -> Get Data -> From Other Sources -> Blank Query ->
// Home -> Advanced Editor. Paste the whole file. Adjust the parameters
// (BaseUrl, ApiKey, PageSize, DataPath, NextPagePath) for your service.
// ============================================================================

let
    // --------------------- CONFIGURE THESE --------------------
    BaseUrl        = "https://api.yourservice.com",
    Endpoint       = "/v1/opportunities",
    ApiKey         = "REPLACE-WITH-KEY-FROM-KEYVAULT",
    PageSize       = 100,
    // JSON path where the list of records lives in the response ("" for root):
    DataPath       = "results",
    // JSON path pointing to the next cursor / offset / URL:
    NextPagePath   = "paging.next.after",
    // How to pass the cursor back: "query_after" | "query_page" | "full_url"
    PaginationMode = "query_after",
    // ----------------------------------------------------------

    // Build the first-page URL
    BuildUrl = (cursor as nullable any) as text =>
        let
            query =
                if PaginationMode = "query_page" then
                    "?limit=" & Text.From(PageSize) &
                    (if cursor <> null then "&page=" & Text.From(cursor) else "&page=1")
                else if PaginationMode = "query_after" then
                    "?limit=" & Text.From(PageSize) &
                    (if cursor <> null then "&after=" & Text.From(cursor) else "")
                else
                    "",
            fullUrl =
                if PaginationMode = "full_url" and cursor <> null then
                    Text.From(cursor)
                else
                    BaseUrl & Endpoint & query
        in
            fullUrl,

    Fetch = (url as text) =>
        let
            raw = Web.Contents(url, [
                Headers = [
                    #"Authorization" = "Bearer " & ApiKey,
                    #"Accept"        = "application/json"
                ],
                ManualStatusHandling = {429, 500, 502, 503, 504}
            ]),
            body = Json.Document(raw)
        in
            body,

    // Walk a JSON path string like "foo.bar.baz"
    PickPath = (record as any, path as text) =>
        if path = "" or path = null then record
        else
            let
                parts = Text.Split(path, "."),
                walk  = List.Accumulate(parts, record, (acc, seg) =>
                    if acc = null then null
                    else try Record.Field(acc, seg) otherwise null)
            in walk,

    // Iterative pager
    PagedList = List.Generate(
        () => [cursor = null, page = Fetch(BuildUrl(null)), done = false],
        each not [done],
        each
            let
                records   = PickPath([page], DataPath),
                nextCur   = PickPath([page], NextPagePath),
                completed = nextCur = null or (records is list and List.Count(records) = 0)
            in
                [
                    cursor = nextCur,
                    page   = if completed then null else Fetch(BuildUrl(nextCur)),
                    done   = completed
                ],
        each PickPath([page], DataPath)
    ),

    // Flatten
    AllRecords     = List.Combine(List.RemoveNulls(PagedList)),
    AsTable        = Table.FromList(AllRecords, Splitter.SplitByNothing(), null, null, ExtraValues.Ignore),
    RecordsTable   =
        if List.Count(AllRecords) > 0 then
            Table.ExpandRecordColumn(
                Table.RenameColumns(AsTable, {{"Column1", "Record"}}),
                "Record",
                Record.FieldNames(AllRecords{0}),
                Record.FieldNames(AllRecords{0})
            )
        else
            AsTable
in
    RecordsTable
