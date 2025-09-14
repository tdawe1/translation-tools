# Translation Pipeline Frontend

A simple web interface for the document translation pipeline that handles PowerPoint presentations and PDF files.

## Features

- **File Upload**: Drag-and-drop or click to upload .pptx and .pdf files
- **Real-time Job Monitoring**: Track translation progress with live updates
- **Model Selection**: Choose between GPT-4o, GPT-4o Mini, and GPT-5
- **Job Management**: View all recent jobs and their status
- **Download Results**: Directly download translated documents
- **Responsive Design**: Works seamlessly on desktop and mobile devices

## Tech Stack

- **Framework**: Next.js 14+ with App Router
- **Language**: TypeScript
- **Styling**: Tailwind CSS + shadcn/ui components
- **State Management**: React Query for server state
- **Forms**: React Hook Form + Zod validation
- **Icons**: Lucide React
- **HTTP Client**: Axios

## Getting Started

### Prerequisites

- Node.js 18+
- npm or yarn
- Backend API running at http://localhost:8000

### Installation

```bash
cd frontend
npm install
```

### Development

```bash
npm run dev
```

The frontend will be available at http://localhost:3000 (or the next available port)

### Build for Production

```bash
npm run build
npm start
```

## Configuration

The frontend expects the backend API to be running. You can configure the API URL in `.env.local`:

```env
NEXT_PUBLIC_API_URL=http://localhost:8000
```

## Project Structure

```
frontend/
├── src/
│   ├── app/
│   │   ├── layout.tsx         # Root layout
│   │   ├── page.tsx           # Home page
│   │   └── globals.css        # Global styles
│   ├── components/
│   │   ├── FileUpload.tsx     # File upload component
│   │   ├── TranslationForm.tsx # Translation settings form
│   │   ├── JobStatus.tsx      # Job status tracking
│   │   └── JobsList.tsx       # Recent jobs list
│   └── lib/
│       └── api.ts             # API client utilities
├── public/                    # Static assets
└── ...
```

## API Integration

The frontend connects to the following endpoints:

- `POST /upload` - Upload files for translation
- `POST /translate` - Start a translation job
- `GET /jobs/{id}` - Get job status
- `GET /jobs` - List all jobs
- `GET /jobs/{id}/download` - Download translated file

See `src/lib/api.ts` for the complete API client implementation.

## Components

- **FileUpload** - Handles file uploads with drag-and-drop
- **TranslationForm** - Model selection and translation initiation
- **JobStatus** - Real-time job progress tracking
- **JobsList** - Shows recent translation jobs

## Styling

The frontend uses Tailwind CSS with a clean, modern design. The color scheme is based on indigo as the primary color with neutral grays for backgrounds and text.