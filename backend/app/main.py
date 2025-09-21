from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from contextlib import asynccontextmanager
import os
import logging
from pathlib import Path

# Import core components
from .core.config import settings
from .database.database import Base, engine
from .api import auth, translate, jobs, sse

# Configure logging
logging.basicConfig(level=settings.LOG_LEVEL)
logger = logging.getLogger(__name__)

# Create tables on startup
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan manager"""
    logger.info("Starting Translation Pipeline API")

    # Create database tables
    Base.metadata.create_all(bind=engine)

    # Ensure directories exist
    os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
    os.makedirs(settings.OUTPUT_DIR, exist_ok=True)

    logger.info(f"Application configuration: {settings.get_environment_info()}")
    yield
    logger.info("Shutting down Translation Pipeline API")

# Initialize FastAPI app
app = FastAPI(
    title=settings.APP_NAME,
    version=settings.VERSION,
    description="A production-ready Japanese-to-English document translation pipeline",
    lifespan=lifespan
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def logging_middleware(request: Request, call_next):
    """Request logging middleware"""
    logger.info(f"{request.method} {request.url}")
    response = await call_next(request)
    logger.info(f"Response status: {response.status_code}")
    return response

# Include API routers
app.include_router(auth.router, prefix="/api/auth", tags=["auth"])
app.include_router(translate.router, prefix="/api", tags=["translate"])
app.include_router(jobs.router, prefix="/api", tags=["jobs"])
app.include_router(sse.router, prefix="/api", tags=["sse"])

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "version": settings.VERSION,
        "name": settings.APP_NAME,
        "openai_configured": settings.is_openai_configured(),
        "redis_configured": settings.is_redis_configured()
    }

@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "Translation Pipeline API",
        "version": settings.VERSION,
        "docs": "/docs",
        "health": "/health"
    }

@app.exception_handler(404)
async def not_found_handler(request: Request, exc):
    return JSONResponse(
        status_code=404,
        content={"detail": "Endpoint not found. Check /docs for available endpoints."}
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        reload=settings.DEBUG,
        log_level=settings.LOG_LEVEL.lower()
    )