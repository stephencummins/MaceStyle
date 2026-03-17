# Mace Style Validator - Technical Architecture

## System Overview

The Mace Style Validator is an automated document validation system that enforces the Mace Control Centre Writing Style Guide using Azure Functions and SharePoint.

> **Note:** Claude AI integration is currently disabled via the `ENABLE_CLAUDE_AI` flag in `config.py`. All hard-coded validation rules continue to run. To re-enable AI-powered style corrections, set `ENABLE_CLAUDE_AI = True`.

## Architecture Diagram

```mermaid
graph TB
    subgraph "SharePoint Online"
        DOC[Document Library]
        RULES[Style Rules List]
        RESULTS[Validation Results List]
        REPORT[HTML Reports]
    end

    subgraph "Orchestration"
        FLOW[Logic App / Power Automate]
    end

    subgraph "Azure"
        FUNC[Azure Function<br/>ValidateDocument]
        GRAPH[Microsoft Graph API]
    end

    subgraph "AI Services"
        CLAUDE[Claude AI API<br/>Anthropic]
    end

    subgraph "Azure AD"
        AUTH[App Registration<br/>Authentication]
    end

    DOC -->|File Created/Modified| FLOW
    FLOW -->|HTTP POST<br/>File Content + Metadata| FUNC

    FUNC -->|Authenticate| AUTH
    AUTH -->|Access Token| FUNC

    FUNC -->|Fetch Style Rules| GRAPH
    GRAPH -->|Read| RULES

    FUNC -->|Download File| GRAPH
    GRAPH -->|Read| DOC

    FUNC -->|Validate Text| CLAUDE
    CLAUDE -->|Corrections| FUNC

    FUNC -->|Upload Fixed File| GRAPH
    GRAPH -->|Write| DOC

    FUNC -->|Upload HTML Report| GRAPH
    GRAPH -->|Write| REPORT

    FUNC -->|Save Validation Result| GRAPH
    GRAPH -->|Create Item| RESULTS

    FUNC -->|Update Metadata| GRAPH
    GRAPH -->|Update| DOC

    FUNC -->|Response| FLOW
    FLOW -->|Update Status| DOC

    style DOC fill:#e3f2fd
    style RULES fill:#e3f2fd
    style RESULTS fill:#e3f2fd
    style REPORT fill:#e3f2fd
    style FUNC fill:#fff3e0
    style CLAUDE fill:#f3e5f5
    style FLOW fill:#e8f5e9
```

## Component Details

### 1. SharePoint Document Library
**Purpose:** Store documents and trigger validation

**Columns:**
- `ValidationStatus` (Choice): Not Validated, Validate Now, Validating..., Passed, Review Required, Failed
- `ValidationResultLink` (Hyperlink): Link to Validation Results list item

**Triggers:** Logic App (or Power Automate) on file create/modify

---

### 2. SharePoint Style Rules List
**Purpose:** Centralized style rule configuration

**Columns:**
- `Title` (Text): Rule description
- `RuleType` (Choice): Font, Language, Grammar, Punctuation, etc.
- `DocumentType` (Choice): Word, Visio, Excel, PowerPoint, Both, All
- `CheckValue` (Text): What to check
- `ExpectedValue` (Text): Correct value
- `AutoFix` (Yes/No): Can be auto-corrected
- `UseAI` (Yes/No): Use Claude AI for validation
- `Priority` (Number): Execution order

**Examples:**
- "Use British English spelling - 'finalised' not 'finalized'"
- "All text must use Arial font"
- "No contractions - use 'cannot' not 'can't'"

---

### 3. SharePoint Validation Results List
**Purpose:** Track validation history and results

**Columns:**
- `Title` (Text): "Validation: {filename}"
- `FileName` (Text): Document name
- `ValidationDate` (DateTime): When validated
- `Status` (Choice): Passed, Review Required, Failed
- `IssuesFound` (Text): Count of issues
- `IssuesFixed` (Text): Count of fixes
- `ReportLink` (Hyperlink): Link to HTML report

**Features:**
- Permanent history of all validations
- Links to detailed HTML reports
- Searchable and filterable

---

### 4. Logic App (or Power Automate)
**Purpose:** Orchestrate validation workflow

**Deployment:** ARM template at `infra/logic-app.json` (preferred for production). Power Automate can also be used for dev/testing.

**Trigger:** When a file is created or modified (SharePoint)

**Actions:**
1. Filter to supported file types (.docx, .xlsx, .pptx, .vsdx, etc.)
2. Set ValidationStatus to "Validating..."
3. Get file properties and content (base64 encode)
4. HTTP POST to Azure Function
5. Parse JSON response
6. If fixes applied, upload corrected file back to SharePoint
7. Update document metadata (status, description, report link, etc.)

**Data Flow:**
```json
Request to Azure Function:
{
  "itemId": 123,
  "fileName": "document.docx",
  "fileContent": "base64...",
  "fileUrl": "/sites/Site/Shared Documents/document.docx"
}

Response from Azure Function:
{
  "status": "Passed",
  "issuesFound": 10,
  "issuesFixed": 10,
  "reportUrl": "https://...",
  "validationResultUrl": "https://...",
  "reportLink": {
    "Description": "View HTML Report",
    "Url": "https://..."
  },
  "validationResultLink": {
    "Description": "View Validation Result",
    "Url": "https://..."
  },
  "fixedFileContent": "base64..."
}
```

---

### 5. Azure Function (ValidateDocument)
**Purpose:** Core validation logic

**Runtime:** Python 3.11

**Key Dependencies:**
- `python-docx`: Word document manipulation
- `python-pptx`: PowerPoint document manipulation
- `openpyxl`: Excel document manipulation
- `vsdx`: Visio document manipulation
- `anthropic`: Claude AI SDK (Word only)
- `msal`: Microsoft authentication
- `requests`: HTTP client

**Validation Flow:**
```mermaid
flowchart TD
    START([HTTP Request]) --> AUTH[Authenticate with Graph API]
    AUTH --> STATUS1[Update Status: Validating...]
    STATUS1 --> RULES[Fetch Style Rules]
    RULES --> DOWNLOAD[Download Document]
    DOWNLOAD --> VALIDATE[Validate Document]

    VALIDATE --> AI{UseAI Rules?}
    AI -->|Yes| CLAUDE[Send to Claude AI]
    CLAUDE --> APPLY_AI[Apply AI Corrections]

    AI -->|No| HARD[Apply Hard-coded Rules]
    APPLY_AI --> HARD

    HARD --> FONT[Font Fixes]
    FONT --> SAVE[Save Fixed Document]

    SAVE --> UPLOAD{Fixes Applied?}
    UPLOAD -->|Yes| UP_DOC[Upload Fixed Document]
    UPLOAD -->|No| REPORT
    UP_DOC --> REPORT[Generate HTML Report]

    REPORT --> UP_REPORT[Upload HTML Report]
    UP_REPORT --> RES[Save to Validation Results]
    RES --> META[Update Document Metadata]
    META --> STATUS2[Update Status: Passed/Review Required/Failed]
    STATUS2 --> END([Return Response])

    style CLAUDE fill:#f3e5f5
    style UPLOAD fill:#fff3e0
    style STATUS2 fill:#e8f5e9
```

**Key Modules:**
- `word_validator.py`: Word (.docx) validation with AI + hard-coded rules
- `visio_validator.py`: Visio (.vsdx) validation — hard-coded rules only
- `excel_validator.py`: Excel (.xlsx) validation — hard-coded rules only
- `powerpoint_validator.py`: PowerPoint (.pptx) validation — hard-coded rules only
- `ai_client.py`: Claude AI integration (Word only)
- `sharepoint_client.py`: Graph API operations
- `report.py`: HTML report generation (summary + collapsible before/after diffs)
- `sharepoint_results.py`: Validation Results list operations

**Note:** AI validation (Claude) is enabled for Word documents only. Visio, Excel, and PowerPoint use hard-coded rules only — AI was disabled for these formats because diagram/spreadsheet/slide text produces too many false positives and unreliable text write-back.

---

### 6. Claude AI Integration (Currently Disabled)
**Purpose:** Advanced language validation for Word documents

> **Status:** Disabled. Set `ENABLE_CLAUDE_AI = True` in `config.py` to re-enable.

**Model:** Claude Haiku 4.5 (fast, cost-effective)

**Capabilities:**
- British English spelling corrections
- Contraction expansion
- Symbol replacement (& → and, % → percent)
- Grammar improvements
- Number formatting

**Process:**
1. Extract all text from document
2. Build dynamic prompt from SharePoint rules (UseAI=True)
3. Send to Claude API
4. Parse JSON response with corrections
5. Apply corrections paragraph by paragraph

**Prompt Structure:**
```
You are a professional document editor applying the Mace Control Centre Writing Style Guide.

Apply ALL of the following corrections:
- Use British English spelling - 'finalised' not 'finalized'
- Use British English spelling - 'colour' not 'color'
- No contractions - use 'cannot' not 'can't'
- Avoid ampersand (&) - use 'and' instead

Return JSON: {"corrected_text": "...", "changes_made": 5}
```

---

### 7. Microsoft Graph API
**Purpose:** SharePoint data access

**Permissions Required:**
- `Sites.ReadWrite.All`: Read/write SharePoint content
- `Files.ReadWrite.All`: Read/write documents

**Key Operations:**
- Read/write files in document libraries
- Read/write list items
- Update file metadata
- Query list data

**Authentication:**
- MSAL (Microsoft Authentication Library)
- Client credentials flow (App registration)

---

### 8. Azure App Registration
**Purpose:** Secure authentication

**Configuration:**
- **Tenant ID**: Azure AD tenant
- **Client ID**: Application ID
- **Client Secret**: Authentication credential

**Permissions:**
- Microsoft Graph API delegated permissions
- SharePoint scopes

**Security:**
- Credentials stored in Azure Function App Settings
- Never exposed in code or logs

---

## Data Flow Sequence

```mermaid
sequenceDiagram
    participant User
    participant SP as SharePoint
    participant PA as Logic App
    participant AF as Azure Function
    participant Graph as Graph API
    participant Claude as Claude AI

    User->>SP: Upload/Modify Document
    SP->>PA: Trigger: File Changed
    PA->>SP: Get File Content
    SP-->>PA: Base64 Content
    PA->>AF: POST /api/validatedocument

    AF->>Graph: Authenticate
    Graph-->>AF: Access Token

    AF->>Graph: Fetch Style Rules
    Graph-->>AF: Rules List

    AF->>Claude: Validate Text (AI Rules)
    Claude-->>AF: Corrections

    AF->>AF: Apply Hard-coded Rules
    AF->>AF: Generate HTML Report

    AF->>Graph: Upload Fixed Document
    Graph->>SP: Update File

    AF->>Graph: Upload HTML Report
    Graph->>SP: Create Report File

    AF->>Graph: Create Validation Result
    Graph->>SP: Add to Results List

    AF->>Graph: Update Document Metadata
    Graph->>SP: Add Result Link

    AF-->>PA: JSON Response
    PA->>SP: Update Status Column
    SP-->>User: Validation Complete
```

---

## Scalability & Performance

### Performance Characteristics
- **Average processing time**: 5-15 seconds per document
- **Bottlenecks**: Claude AI API calls (2-5 seconds)
- **Concurrency**: Azure Functions auto-scale

### Optimization Strategies
1. **Single AI call**: All AI rules processed in one request
2. **Parallel operations**: Font fixes during AI processing
3. **Efficient Graph API**: Batch operations where possible
4. **Caching**: Graph API tokens cached for 60 minutes

### Cost Considerations
- **Azure Functions**: Consumption plan (pay-per-execution)
- **Claude AI**: ~$0.01 per document (Haiku model)
- **Storage**: SharePoint included in Microsoft 365

---

## Security Architecture

### Authentication Flow
```mermaid
flowchart LR
    AF[Azure Function] -->|Client Credentials| AAD[Azure AD]
    AAD -->|JWT Token| AF
    AF -->|Bearer Token| GRAPH[Graph API]
    GRAPH -->|Authorized| SP[SharePoint]

    style AAD fill:#f3e5f5
    style AF fill:#fff3e0
```

### Security Measures
1. **App Registration**: Service principal with minimal permissions
2. **Secret Management**: Azure Key Vault / App Settings (encrypted)
3. **Network Security**: HTTPS only, Azure Function CORS policies
4. **Data Privacy**: Documents processed in-memory, not persisted
5. **Audit Trail**: Validation Results list tracks all operations

---

## Monitoring & Logging

### Azure Function Logging
- **Application Insights**: Performance metrics, errors
- **Log Stream**: Real-time debugging
- **Metrics**: Request count, duration, failures

### Key Metrics to Monitor
- Validation success rate
- Average processing time
- Claude API errors
- SharePoint API throttling
- Document upload failures

### Alerting
- Failed validations > 10%
- Average processing time > 30 seconds
- API errors > 5%

---

## Technology Stack Summary

| Layer | Technology | Purpose |
|-------|-----------|---------|
| **Frontend** | SharePoint Online | Document storage, UI |
| **Workflow** | Logic App (ARM template) | Orchestration |
| **Backend** | Azure Functions (Python 3.11) | Validation logic |
| **AI** | Claude Haiku 4.5 (Anthropic) | Language processing (currently disabled) |
| **API** | Microsoft Graph API | SharePoint integration |
| **Auth** | Azure AD / MSAL | Authentication |
| **Storage** | SharePoint Lists & Libraries | Data persistence |

---

## Version Information

- **Current Version**: v5.1
- **Last Updated**: March 2026
- **Python Version**: 3.11
- **Azure Functions Runtime**: 4.x
