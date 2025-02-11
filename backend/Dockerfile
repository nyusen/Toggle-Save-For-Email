FROM python:3.11-slim AS local
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

RUN apt-get update 
RUN apt-get install libgomp1 && apt-get install -y curl && apt-get install -y git
RUN curl -sL https://aka.ms/InstallAzureCLIDeb | bash

RUN curl https://packages.microsoft.com/keys/microsoft.asc | tee /etc/apt/trusted.gpg.d/microsoft.asc
RUN curl https://packages.microsoft.com/config/debian/12/prod.list | tee /etc/apt/sources.list.d/mssql-release.list
RUN sed -i 's|signed-by=/usr/share/keyrings/microsoft-prod.gpg||g' /etc/apt/sources.list.d/mssql-release.list
RUN apt-get update
RUN ACCEPT_EULA=Y apt-get install -y msodbcsql18
RUN echo 'export PATH="$PATH:/opt/mssql-tools18/bin"' >> ~/.bashrc
RUN apt-get install -y unixodbc-dev

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


FROM python:3.11-slim AS deployment
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

RUN apt-get update 
RUN apt-get install libgomp1 && apt-get install -y curl && apt-get install -y git
RUN curl -sL https://aka.ms/InstallAzureCLIDeb | bash

RUN curl https://packages.microsoft.com/keys/microsoft.asc | tee /etc/apt/trusted.gpg.d/microsoft.asc
RUN curl https://packages.microsoft.com/config/debian/12/prod.list | tee /etc/apt/sources.list.d/mssql-release.list
RUN sed -i 's|signed-by=/usr/share/keyrings/microsoft-prod.gpg||g' /etc/apt/sources.list.d/mssql-release.list
RUN apt-get update
RUN ACCEPT_EULA=Y apt-get install -y msodbcsql18
RUN echo 'export PATH="$PATH:/opt/mssql-tools18/bin"' >> ~/.bashrc
RUN apt-get install -y unixodbc-dev

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

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8502"]