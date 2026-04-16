from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from backend.routes import browse, export, health, master, match, registry, sf_issues, titles, workflow

app = FastAPI(title="crewdbupdate")

app.include_router(health.router, prefix="/api", tags=["health"])
app.include_router(titles.router, prefix="/api", tags=["titles"])
app.include_router(export.router, prefix="/api", tags=["export"])
app.include_router(master.router, prefix="/api", tags=["master"])
app.include_router(match.router, prefix="/api", tags=["match"])
app.include_router(browse.router, prefix="/api", tags=["browse"])
app.include_router(workflow.router, prefix="/api", tags=["workflow"])
app.include_router(registry.router, prefix="/api", tags=["registry"])
app.include_router(sf_issues.router, prefix="/api", tags=["sf_issues"])

frontend_dir = Path(__file__).resolve().parents[1] / "frontend"


@app.get("/")
def root_redirect() -> RedirectResponse:
    return RedirectResponse(url="/export", status_code=307)


@app.get("/export")
def export_page() -> FileResponse:
    return FileResponse(frontend_dir / "export.html")


@app.get("/titles")
def titles_page() -> FileResponse:
    return FileResponse(frontend_dir / "titles.html")


@app.get("/match")
def match_page() -> FileResponse:
    return FileResponse(frontend_dir / "match.html")


@app.get("/sf_issues")
def sf_issues_page() -> FileResponse:
    return FileResponse(frontend_dir / "sf_issues.html")


@app.get("/sf_browser")
def sf_browser_page() -> FileResponse:
    return FileResponse(frontend_dir / "sf_browser.html")


@app.get("/crew_explorer")
def crew_explorer_page() -> FileResponse:
    return FileResponse(frontend_dir / "crew_explorer.html")


@app.get("/combined_browser")
def combined_browser_redirect() -> RedirectResponse:
    return RedirectResponse(url="/crew_explorer", status_code=301)


@app.get("/registry")
def registry_page() -> FileResponse:
    return FileResponse(frontend_dir / "registry.html")


@app.get("/names_browser")
def names_browser_redirect() -> RedirectResponse:
    return RedirectResponse(url="/registry", status_code=301)


app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
