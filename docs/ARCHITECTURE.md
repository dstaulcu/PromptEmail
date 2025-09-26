# PromptEmail Architecture Guide

## System Overview

PromptEmail is an Office Add-in that integrates AI-powered email analysis with secure telemetry collection. The system consists of multiple interconnected components spanning client-side Office.js integration, serverless AWS infrastructure, and enterprise analytics.

## High-Level Architecture

```mermaid
graph TB
    subgraph "Client Environment"
        OL[Outlook Desktop]
        OA[Office Add-in]
        OJS[Office.js API]
    end
    
    subgraph "Static Hosting"
        S3[AWS S3 Bucket]
        MS[Microsoft Office.js CDN]
    end
    
    subgraph "AI Providers"
        OLLAMA[Ollama Local]
        ONSITE[OnSite AI APIs]
        OPENAI[OpenAI Compatible]
    end
    
    subgraph "AWS Infrastructure"
        APIG[API Gateway]
        LAMBDA[Lambda Function]
        CW[CloudWatch Logs]
    end
    
    subgraph "Analytics Platform"
        SPLUNK[Splunk Enterprise]
        DASH[Dashboards]
    end
    
    OL --> OA
    OA --> OJS
    OA --> S3
    OJS --> MS
    
    OA --> OLLAMA
    OA --> ONSITE
    OA --> OPENAI
    
    OA --> APIG
    APIG --> LAMBDA
    LAMBDA --> SPLUNK
    LAMBDA --> CW
    
    SPLUNK --> DASH
```

## Component Interaction Flows

### 1. Add-in Initialization Sequence

```mermaid
sequenceDiagram
    participant U as User
    participant O as Outlook
    participant A as Add-in
    participant S3 as S3 Bucket
    participant MS as Microsoft CDN
    
    U->>O: Opens Outlook
    O->>A: Loads Add-in Manifest
    A->>S3: Fetch HTML/CSS/JS Assets
    S3->>A: Return Static Assets
    A->>MS: Load Office.js Library
    MS->>A: Return Office.js
    A->>O: Initialize Office API
    O->>A: Ready Event
    A->>U: Display Interface
```

### 2. Email Analysis Workflow

```mermaid
sequenceDiagram
    participant U as User
    participant A as Add-in
    participant O as Office API
    participant AI as AI Provider
    participant T as Telemetry
    
    U->>A: Click "Analyze Email"
    A->>O: Office.context.mailbox.item.body.getAsync()
    O->>A: Return Email Content
    A->>A: Process & Sanitize Content
    A->>AI: Send Analysis Request
    AI->>A: Return AI Response
    A->>T: Log Analysis Event
    A->>U: Display Results
    
    Note over A,T: Telemetry includes performance metrics,<br/>user behavior, and error tracking
```

### 3. Telemetry Pipeline Architecture

```mermaid
sequenceDiagram
    participant A as Add-in Client
    participant AG as API Gateway
    participant L as Lambda Function
    participant S as Splunk HEC
    participant CW as CloudWatch
    
    A->>A: Queue Telemetry Events
    A->>AG: POST /prod/events (Batch)
    AG->>L: Invoke Lambda
    L->>L: Validate & Transform Events
    L->>S: Forward to Splunk HEC
    S->>L: Acknowledgment
    L->>CW: Log Processing Details
    L->>AG: Return Success
    AG->>A: HTTP 200 Response
    
    Note over A,S: Events include session data,<br/>performance metrics, errors,<br/>and user behavior analytics
```

### 4. Deployment Pipeline Flow

```mermaid
sequenceDiagram
    participant D as Developer
    participant W as Webpack
    participant S as Deploy Script
    participant S3 as S3 Bucket
    participant CF as CloudFormation
    participant L as Lambda
    
    D->>W: npm run build
    W->>W: Bundle Assets
    W->>S: Built Assets Ready
    S->>S: Process Manifest & URLs
    S->>S3: Upload Static Assets
    S->>CF: Deploy Infrastructure
    CF->>L: Update Lambda Code
    L->>L: Configure Environment
    S->>D: Deployment Complete
```

## Environment & Deployment Architecture

## Technology Stack

| Layer | Technology | Purpose |
|-------|------------|---------|
| **Frontend** | Office.js, JavaScript ES6+, HTML5, CSS3 | Client-side add-in interface |
| **Build System** | Webpack 5, Babel, npm scripts | Asset bundling and build automation |
| **Static Hosting** | AWS S3, CloudFront CDN (optional)| Static asset delivery and caching |
| **Serverless API** | AWS API Gateway, Lambda (Node.js) | Telemetry collection and processing |
| **Analytics** | Splunk Enterprise, HTTP Event Collector | Data analysis and visualization |
| **Infrastructure** | AWS CloudFormation, PowerShell scripts | Infrastructure as Code |
| **AI Integration** | OpenAI API, Ollama, Custom APIs | AI-powered email analysis and style personalization |
| **Settings Storage** | Office.js RoamingSettings + localStorage | Dual-layer user preferences and writing samples persistence |
| **Monitoring** | AWS CloudWatch, Splunk dashboards | System monitoring and alerting |

## Data Flow Patterns

### Event-Driven Telemetry

```mermaid
stateDiagram-v2
    [*] --> SessionStart
    SessionStart --> EmailSelected
    EmailSelected --> AnalysisRequested
    AnalysisRequested --> AIProcessing
    AIProcessing --> ResultsDisplayed
    ResultsDisplayed --> UserInteraction
    UserInteraction --> EmailSelected : Next Email
    UserInteraction --> SessionEnd
    SessionEnd --> [*]
    
    SessionStart : Log Session Start
    EmailSelected : Log Email Context
    AnalysisRequested : Log Analysis Request
    AIProcessing : Log Performance Metrics
    ResultsDisplayed : Log Results & Timing
    UserInteraction : Log User Actions
    SessionEnd : Log Session Summary
```

### Writing Samples & Style Personalization

The system includes a sophisticated writing samples feature that allows users to train the AI on their personal writing style for more authentic email responses.

```mermaid
graph TB
    subgraph "User Interface"
        WSI[Writing Samples Input]
        WSM[Sample Management UI]
        SS[Style Settings]
    end
    
    subgraph "Settings Management"
        SM[SettingsManager]
        WS[Writing Samples Array]
        SA[Style Analysis Settings]
        ST[Style Strength Config]
    end
    
    subgraph "AI Integration"
        AI[AIService]
        PP[Prompt Processing]
        SF[Sample Filtering]
        SG[Style-Aware Generation]
    end
    
    subgraph "Data Storage"
        OS[Office.js Settings API]
        RS[Roaming Settings]
        LS[Local Storage Backup]
    end
    
    WSI --> SM
    WSM --> SM
    SS --> SM
    
    SM --> WS
    SM --> SA
    SM --> ST
    
    SM --> OS
    OS --> RS
    OS --> LS
    
    SM --> AI
    AI --> PP
    PP --> SF
    SF --> SG
    
    WS --> SF
    SA --> SF
    ST --> SF
```

#### Writing Samples Data Flow

1. **Sample Collection**: Users input examples of their written communication
2. **Storage**: Samples stored with metadata (date, word count, unique ID) using Office.js Settings API
3. **Style Analysis**: System analyzes writing patterns when style-analysis is enabled
4. **Prompt Enhancement**: AI prompts include selected samples based on style strength settings
5. **Contextual Generation**: AI generates responses that match user's writing style and tone

#### Style Strength Configuration

| Setting | Sample Count | Use Case |
|---------|-------------|-----------|
| **Light** | 2 samples | Subtle style influence |
| **Medium** | 3 samples | Balanced personalization |
| **Strong** | 5 samples | Maximum style matching |

### Security & Privacy Considerations

**Data Protection:**
- Writing samples stored using dual-layer approach: Office.js RoamingSettings (primary) with localStorage fallback
- Cross-device roaming when Office.js RoamingSettings available, local-only when using localStorage fallback
- No external transmission of samples except to user-configured AI providers
- Automatic encryption through Office 365 when using RoamingSettings, browser security when using localStorage
- User control over sample retention and deletion

**Token Management:**
- Intelligent sample selection to optimize context window usage
- Monitoring and logging of token consumption for transparency
- Fallback mechanisms when samples approach token limits
- Debug logging for token usage analysis

---

## Additional Resources

- [Developer Guide](docs/DEVELOPER_GUIDE.md) - Development setup and workflows
- [Deployment Guide](tools/README.md) - Infrastructure deployment
