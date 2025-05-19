# Use a Python image with uv pre-installed for fast, reproducible builds
FROM ghcr.io/astral-sh/uv:python3.12-bookworm-slim AS uv

WORKDIR /app

# Enable bytecode compilation
ENV UV_COMPILE_BYTECODE=1
ENV UV_LINK_MODE=copy

# Install dependencies using uv (if you have pyproject.toml and uv.lock)
# If you only have requirements.txt, you can skip this block and use pip in the final image
# RUN --mount=type=cache,target=/root/.cache/uv \
#     --mount=type=bind,source=uv.lock,target=uv.lock \
#     --mount=type=bind,source=pyproject.toml,target=pyproject.toml \
#     uv sync --frozen --no-install-project --no-dev --no-editable

# If you use requirements.txt, copy and install here
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Add the rest of the project source code
ADD . /app

FROM python:3.12-slim-bookworm

WORKDIR /app

# Copy installed dependencies from the builder
COPY --from=uv /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=uv /usr/local/bin /usr/local/bin

# Copy your app code
COPY . /app

# Set environment variable for Excel directory (can be overridden at runtime)
ENV EXCEL_FILES_DIR=/app/excel_files

# Create the directory (optional, can be omitted if you want the client to provide it)
# RUN mkdir -p /app/excel_files

# Expose the port (if using HTTP)
EXPOSE 8000

# Set the entrypoint to your server (main.py or advanced_server.py)
# CMD ["python", "main.py"]
# Or for advanced server:
CMD ["python", "advanced_server.py"]