FROM python:3.12-slim

LABEL org.opencontainers.image.title="Word AI"
LABEL org.opencontainers.image.description="Structure-preserving Word DOCX editing MCP server and Office.js bridge for AI agents"
LABEL org.opencontainers.image.source="https://github.com/flyfish-dev/word-ai"
LABEL org.opencontainers.image.licenses="AGPL-3.0-or-later"
LABEL io.modelcontextprotocol.server.name="io.github.flyfish-dev/word-ai"

ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY pyproject.toml README.md LICENSE requirements.txt ./
COPY schemas ./schemas
COPY word_ai_mcp ./word_ai_mcp

RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir .

WORKDIR /workspace

ENTRYPOINT ["python", "-m", "word_ai_mcp.server"]
CMD ["--root", "/workspace", "--allow-root", "/documents"]
