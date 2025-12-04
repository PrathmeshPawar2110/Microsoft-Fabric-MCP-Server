# 📘 Microsoft Fabric MCP Server – README

## Overview

This project implements a **Model Context Protocol (MCP)** server for Microsoft Fabric using **FastAPI** and **FastMCP**. It provides an API layer for controlling and automating Fabric resources such as Workspaces, Items, Domains, Pipelines, and Lakehouse data via secure, session-based or API-key access.

---

## 🧩 Key Capabilities

### 1. Workspace Management

- Create, update, delete, and list workspaces
- Assign/unassign workspaces to/from:
  - Capacities
  - Deployment stages
  - Domains (by capacity or workspace ID)
- Manage role assignments for workspaces

### 2. Item Management

- Create, delete, update items (Reports, Dashboards, etc.)
- Get item definitions
- List items with filtering
- Apply/unapply tags to items
- List item connections

### 3. Folder Management

- Create, delete, update, move folders
- List folders (recursively or flat)
- Retrieve folder metadata

### 4. Deployment Pipelines

- Create/delete pipelines
- Create/update stages
- Assign/unassign workspaces to/from stages
- Deploy content across stages with optional notes
- Manage pipeline role assignments
- List pipeline stages, operations, and deployed items

### 5. Domain Management

- Create, update, delete domains
- List domains and their workspaces
- Assign/unassign workspaces to domains
- Bulk assign/unassign roles to principals in a domain
- Unassign all or specific workspaces

### 6. Tagging

- Apply tags to items
- Unapply tags
- List all tags in tenant scope

### 7. Capacities

- List all available Fabric capacities

### 8. Lakehouse SQL Access

- Query lakehouse data using SQL via an authenticated ODBC connection

---

## 🔐 Authentication

- Uses **Microsoft Identity Platform (MSAL)** for authentication
- Redis is used for session storage and API-key mapping
- Users can authenticate via OAuth2 flow or use API keys (with TTL)

---

## ⚙️ Environment Configuration

Set the following variables in a `.env` file:

CLIENT_ID=...

CLIENT_SECRET=...

TENANT_ID=...

SESSION_SECRET_KEY=...

APP_URL=[https://your-app-url]()

REDIS_HOST=your-redis-host

REDIS_PORT=6380

REDIS_KEY=your-redis-password

## 🚀 Running the Server

**Navigate to source directory**

```pip
cd ./fabric_mcp_server_user
```

**Install dependencies:**

```pip
pip install -r requirements.txt
```

**Run the FastAPI app:**

```
uvicorn main:app --reload
```

## 🛠️ Extending the MCP

You can expose new capabilities by wrapping functions with `@mcp.tool()` and following the pattern already defined.
