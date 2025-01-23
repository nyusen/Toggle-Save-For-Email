FROM python:3.11-slim AS local
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

RUN apt-get update 
RUN apt-get install libgomp1 && apt-get install -y curl && apt-get install -y git
RUN curl -sL https://aka.ms/InstallAzureCLIDeb | bash

WORKDIR /code

COPY ./pyproject.toml ./pyproject.toml 
COPY ./uv.lock ./uv.lock 

RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --inexact --no-dev

# Place executables in the environment at the front of the path
ENV PATH="/code/.venv/bin:$PATH" \
    UV_COMPILE_BYTECODE=1 \
    VIRTUAL_ENV=/code/.venv

# Copy the rest of the application
COPY ./app ./app

# Expose the port the app runs on
EXPOSE 8000

# Command to run the application
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
