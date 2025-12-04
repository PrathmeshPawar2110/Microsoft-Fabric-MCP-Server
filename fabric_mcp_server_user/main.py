import json
import logging
import msal
import os
import redis
import requests
import secrets
import threading
import pyodbc
import struct
import json
import uuid
from contextlib import asynccontextmanager
from contextvars import ContextVar
from dotenv import load_dotenv
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import RedirectResponse
from fastmcp import FastMCP
from starlette.middleware.sessions import SessionMiddleware
from typing import Any, Dict, List, Optional
import uvicorn
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
matplotlib.use('Agg')
from pathlib import Path
from datetime import datetime
from azure.identity import ClientSecretCredential

# --- Basic Setup & Configuration ---
logger = logging.getLogger("microsoft-fabric-mcp")
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY")
APP_URL = os.getenv("APP_URL")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_PATH = "/auth/callback"
REDIRECT_URI = f"{APP_URL}{REDIRECT_PATH}"
SCOPE = ["https://api.fabric.microsoft.com/.default"]
REDIS_HOST = os.getenv("REDIS_HOST")
REDIS_PORT = int(os.getenv("REDIS_PORT", 6380))
REDIS_KEY = os.getenv("REDIS_KEY")
SESSION_TTL_SECONDS = 7200


redis_client = redis.Redis(
    host=REDIS_HOST,
    port=REDIS_PORT,
    password=REDIS_KEY,
    ssl=True,
    decode_responses=True
)

# --- In-Memory State Management ---
user_contexts: Dict[str, Dict[str, Any]] = {}
user_contexts_lock = threading.Lock()

api_keys: Dict[str, Dict[str, Any]] = {}
api_keys_lock = threading.Lock()

current_user_context: ContextVar[Optional[str]] = ContextVar('current_user_context', default=None)


class FabricClient:
    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://api.fabric.microsoft.com/v1"
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        })
        

    def create_item(self, workspace_id: str, displayName: str, item_type: str, description: Optional[str] = None, folder_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new item in the specified workspace"""
        try:
            headers = self.session.headers

            payload = {
                "displayName": displayName,
                "type": item_type
            }
            
            if description:
                payload["description"] = description

            if folder_id:
                payload["folderId"] = folder_id

            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/items",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error creating item: {e}")
            raise
    
    
    def delete_item(self, workspace_id: str, item_id: str) -> Dict[str, Any]:
        """Delete a specific item from a workspace"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}",
                headers=headers
            )
            response.raise_for_status()

            if response.status_code == 204 or not response.text.strip():
                return {
                    "success": True,
                    "message": "Item deleted successfully",
                    "workspace_id": workspace_id,
                    "item_id": item_id
                }

            try:
                json_response = response.json()
                if isinstance(json_response, dict):
                    return json_response.get("value", json_response)
                else:
                    return {"success": True, "data": json_response}
            except ValueError:
                return {
                    "success": True,
                    "message": "Item deleted successfully",
                    "item_id": item_id,
                    "workspace_id": workspace_id
                }
        except Exception as e:
            logger.error(f"Error deleting item: {e}")
            raise


    def get_item_definition(self, workspace_id: str, item_id: str, format: Optional[str] = None) -> Dict[str, Any]:
        """Get the definition of a specific item"""
        try:
            headers = self.session.headers
            
            url = f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}/getDefinition"
            if format:
                url += f"?format={format}"
            
            response = self.session.post(
                url,
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting item definition: {e}")
            raise


    def list_item_connections(self, workspace_id: str, item_id: str) -> List[Dict[str, Any]]:
        """List all connections for a specific item"""
        try: 
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}/connections",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing item connections: {e}")
            raise


    def list_items(self, workspace_id: str, item_type: Optional[str] = None, recursive: Optional[bool] = False, root_folder_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """List all items of a specific type in a workspace"""
        try:
            headers = self.session.headers
            url = f"{self.base_url}/workspaces/{workspace_id}/items"
            
            params = {}
            if item_type:
                params['type'] = item_type
            if recursive:
                params['recursive'] = str(recursive).lower()
            if root_folder_id:
                params['rootFolderId'] = root_folder_id
            
            response = self.session.get(
                url,
                headers=headers,
                params=params
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing items: {e}")
            raise


    def update_item(self, workspace_id: str, item_id: str, updatedDisplayName: str, updatedDescription: Optional[str] = None) -> Dict[str, Any]:
        """Update a specific item in a workspace"""
        try:
            headers = self.session.headers
            payload = {
                "displayName": updatedDisplayName
            }
            if updatedDescription:
                payload["description"] = updatedDescription

            response = self.session.patch(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating item: {e}")
            raise


    def add_role_assignment_deployment_pipeline(self, pipeline_id: str, role: str, principal_id: str, principal_type: str) -> Dict[str, Any]:
        """Add a role assignment to a deployment pipeline"""
        try:
            headers = self.session.headers
            
            payload = {
                "principal": {
                    "id": principal_id,
                    "type": principal_type
                },
                "role": role
            }
            
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/roleAssignments",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            if response.status_code == 200:
                return {
                    "success": True,
                    "message": "Role assignment added successfully",
                    "pipeline_id": pipeline_id
                }
            
            if response.text.strip():
                return response.json()
            else:
                return {"success": True, "message": "Operation completed"}
        except Exception as e:
            logger.error(f"Error adding role assignment: {e}")
            raise


    def assign_workspace_to_a_stage_in_deployment_pipeline(self, pipeline_id: str, stage_id: str, workspace_id: str) -> Dict[str, Any]:
        """Assign a workspace to a specific stage in a deployment pipeline"""
        try:
            headers = self.session.headers
            
            payload = {
                "workspaceId": workspace_id
            }
            
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}/assignWorkspace",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            
            if response.status_code == 204 or not response.text.strip():
                return {
                    "success": True,
                    "message": "Workspace assigned successfully",
                    "pipeline_id": pipeline_id,
                    "stage_id": stage_id, 
                    "workspace_id": workspace_id
                }
            
            try:
                json_response = response.json()
                if isinstance(json_response, dict):
                    return json_response.get("value", json_response)
                else:
                    return {"success": True, "data": json_response}
            except ValueError:
                return {
                    "success": True,
                    "message": "Workspace assigned successfully",
                    "pipeline_id": pipeline_id,
                    "stage_id": stage_id,
                    "workspace_id": workspace_id
                }
                
        except Exception as e:
            logger.error(f"Error assigning workspace to deployment pipeline: {e}")
            raise


    def create_deployment_pipeline(self, display_name: str, description: Optional[str] = None, stages: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
        """Create a deployment pipeline"""
        try:
            headers = self.session.headers
            
            if not stages:
                stages = [
                    {"order": 0, "displayName": "Development"},
                    {"order": 1, "displayName": "Test"},
                    {"order": 2, "displayName": "Production"}
                ]
            
            payload = {
                "displayName": display_name,
                "stages": stages
            }
            
            if description:
                payload["description"] = description
            
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error creating deployment pipeline: {e}")
            raise


    def delete_deployment_pipeline(self, pipeline_id: str) -> Dict[str, Any]:
        """Delete a specific deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}",
                headers=headers
            )
            response.raise_for_status()
            return {"success": True}
        except Exception as e:
            logger.error(f"Error deleting deployment pipeline: {e}")
            raise


    def delete_deployment_pipeline_role_assignment(self, pipeline_id: str, principal_id: str):
        """Delete role assignment for a deployment pipeline"""
        try:
            headers = self.session.headers
                        
            response = self.session.delete(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/roleAssignments/{principal_id}",
                headers=headers
            )
            response.raise_for_status()
            if response.status_code == 200:
                return {
                    "success": True,
                    "message": "Role assignment deleted successfully",
                    "pipeline_id": pipeline_id,
                    "principal_id": principal_id
                }
            
            if response.text.strip():
                return response.json()
            else:
                return {"success": True, "message": "Operation completed"}
        except Exception as e:
            logger.error(f"Error adding role assignment: {e}")
            raise


    def deploy_stage_content(self, pipeline_id: str, source_stage_id: str, target_stage_id: str, note: Optional[str] = None, items: Optional[List[Dict]] = None) -> Dict[str, Any]:
        """Deploy content to a specific stage in a deployment pipeline"""
        try:
            # Build the request payload according to Microsoft Fabric API specification
            payload = {
                "sourceStageId": source_stage_id,
                "targetStageId": target_stage_id
            }
            
            # Add optional note if provided
            if note:
                payload["note"] = note
           
            # Add items if provided - fix the field mapping
            if items:
                payload["items"] = [
                    {
                        "sourceItemId": item.get('sourceItemId', item.get('id')),  # Support both field names
                        "itemType": item.get('itemType', item.get('type'))        # Support both field names
                    }
                    for item in items
                ]
            
            # Log the request for debugging
            logger.info(f"Deploying to pipeline {pipeline_id}: {payload}")
               
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/deploy",
                json=payload
            )
            
            # Enhanced error handling
            if response.status_code == 400:
                error_detail = response.json() if response.content else {"message": "Bad Request"}
                logger.error(f"Bad Request (400): {error_detail}")
                raise Exception(f"Bad Request: {error_detail.get('message', 'Invalid request format')}")
            elif response.status_code == 401:
                logger.error("Unauthorized (401): Token may be expired or invalid")
                raise Exception("Unauthorized: Please check your access token")
            elif response.status_code == 403:
                logger.error("Forbidden (403): Insufficient permissions")
                raise Exception("Forbidden: You don't have permission to deploy to this pipeline")
            elif response.status_code == 404:
                logger.error(f"Not Found (404): Pipeline {pipeline_id} not found")
                raise Exception(f"Pipeline {pipeline_id} not found")
            
            response.raise_for_status()
            
            # Handle both JSON response and empty response (for 202 Accepted)
            if response.content:
                try:
                    return response.json()
                except:
                    return {"status": "accepted", "status_code": response.status_code}
            else:
                return {"status": "accepted", "status_code": response.status_code}
                
        except requests.exceptions.RequestException as e:
            logger.error(f"Request error deploying stage content: {e}")
            raise Exception(f"Request failed: {str(e)}")
        except Exception as e:
            logger.error(f"Error deploying stage content: {e}")
            raise


    def get_deployment_pipeline(self, pipeline_id: str) -> Dict[str, Any]:
        """Get details of a specific deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting deployment pipeline details: {e}")
            raise


    def get_deployment_pipeline_stage(self, pipeline_id: str, stage_id: str) -> Dict[str, Any]:
        """Get details of a specific stage in a deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting deployment pipeline stage details: {e}")
            raise


    def list_deployment_pipeline_operations(self, pipeline_id: str) -> List[Dict[str, Any]]:
        """List all operations for a deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/operations",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing deployment pipeline operations: {e}")
            raise


    def list_deployment_pipeline_role_assignments(self, pipeline_id: str) -> List[Dict[str, Any]]:
        """List all role assignments for a deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/roleAssignments",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing role assignments: {e}")
            raise


    def list_deployment_pipeline_stage_items(self, pipeline_id: str, stage_id: str) -> List[Dict[str, Any]]:
        """List all items in a specific stage of a deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}/items",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing deployment pipeline stage items: {e}")
            raise


    def list_deployment_pipeline_stages(self, pipeline_id: str) -> List[Dict[str, Any]]:
        """List stages of a specific deployment pipeline"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error getting deployment pipeline stages: {e}")
            raise


    def list_deployment_pipelines(self) -> List[Dict[str, Any]]:
        """List all deployment pipelines"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing deployment pipelines: {e}")
            raise


    def unassign_workspace_from_a_stage_in_deployment_pipeline(self, pipeline_id: str, stage_id: str) -> Dict[str, Any]:
        """Unassign a workspace from a specific stage in a deployment pipeline"""
        try:
            headers = self.session.headers
            
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}/unassignWorkspace",
                headers=headers
            )
            response.raise_for_status()
            return {"status": "success", "message": "Workspace unassigned successfully"}
        except Exception as e:
            logger.error(f"Error unassigning workspace from deployment pipeline: {e}")
            raise


    def update_deployment_pipeline(self, pipeline_id: str, display_name: Optional[str] = None, description: Optional[str] = None, stages: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
        """Update an existing deployment pipeline"""
        try:
            headers = self.session.headers
            payload = {}
            
            if display_name:
                payload["displayName"] = display_name
            if description:
                payload["description"] = description
            if stages is not None:
                payload["stages"] = stages
            
            response = self.session.patch(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating deployment pipeline: {e}")
            raise


    def update_deployment_pipeline_stage(self, pipeline_id: str, stage_id: str, display_name: Optional[str] = None, description: Optional[str] = None) -> Dict[str, Any]:
        """Update a specific stage in a deployment pipeline"""
        try:
            headers = self.session.headers
            payload = {}
            if display_name:
                payload["displayName"] = display_name
            if description:
                payload["description"] = description

            response = self.session.patch(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating deployment pipeline stage: {e}")
            raise


    def assign_workspace_to_domain_by_capacity(self, domain_id: str, capacity_id: str) -> Dict[str, Any]:
        """Assign a workspace to a domain by capacity"""
        try:
            headers = self.session.headers
            payload = {
                "capacityId": capacity_id
            }
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/assignWorkspacesByCapacities",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error assigning workspace to domain: {e}")
            raise
            

    def assign_workspace_to_domain_by_ids(self, domain_id: str, workspace_ids: List[str]) -> Dict[str, Any]:
        """Assign a workspace to a domain"""
        try:
            headers = self.session.headers
            payload = {
                "workspaceIds": workspace_ids
            }
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/assignWorkspaces",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error assigning workspace to domain: {e}")
            raise


    def create_domain(self, domain_name: str, description: Optional[str] = None, parent_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new domain in Microsoft Fabric"""
        try:
            headers = self.session.headers
            payload = {
                "displayName": domain_name,
                "description": description or ""
            }
            if parent_id:
                payload["parentId"] = parent_id
            response = self.session.post(
                f"{self.base_url}/admin/domains",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error creating domain: {e}")
            raise


    def delete_domain(self, domain_id: str) -> Dict[str, Any]:
        """Delete a domain in Microsoft Fabric"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/domains/{domain_id}",
                headers=headers
            )
            response.raise_for_status()
            return {"status": "success", "message": "Domain deleted successfully"}
        except Exception as e:
            logger.error(f"Error deleting domain: {e}")
            raise


    def get_domain(self, domain_id: str) -> Dict[str, Any]:
        """Get details of a specific domain"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/domains/{domain_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting domain details: {e}")
            raise


    def list_domain_workspaces(self, domain_id: str) -> List[Dict[str, Any]]:
        """List all workspaces in a specific domain"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/domains/{domain_id}/workspaces",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing domain workspaces: {e}")
            raise


    def list_domains(self) -> List[Dict[str, Any]]:
        """List all domains in Microsoft Fabric"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/admin/domains",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing domains: {e}")
            raise


    def domain_bulk_assign_roles(self, domain_id: str, role: str, principals: List[Dict]) -> Dict[str, Any]:
        """Bulk assign roles to principals in a domain"""
        try:
            headers = self.session.headers
            payload = {
                "type": role,
                "principals": [
                    {
                        "id": principal["id"],
                        "type": principal["type"]
                    }
                    for principal in principals
                ]
            }
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/roleAssignments/bulkAssign",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error bulk assigning roles in domain: {e}")
            raise


    def domain_bulk_unassign_roles(self, domain_id: str, role: str, principals: List[Dict]) -> Dict[str, Any]:
        """Bulk unassign roles from principals in a domain"""
        try:
            headers = self.session.headers
            payload = {
                "type": role,
                "principals": [
                    {
                        "id": principal["id"],
                        "type": principal["type"]
                    }
                    for principal in principals
                ]
            }
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/roleAssignments/bulkUnassign",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error bulk unassigning roles in domain: {e}")
            raise


    def domain_unassign_all_workspaces(self, domain_id: str) -> Dict[str, Any]:
        """Unassign all workspaces from a domain"""
        try:
            headers = self.session.headers
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/unassignAllWorkspaces",
                headers=headers
            )
            response.raise_for_status()
            return {"status": "success", "message": "All workspaces unassigned from domain successfully"}
        except Exception as e:
            logger.error(f"Error unassigning all workspaces from domain: {e}")
            raise


    def domain_unassign_workspace_by_ids(self, domain_id: str, workspace_ids: List[str]) -> Dict[str, Any]:
        """Unassign specific workspaces from a domain"""
        try:
            headers = self.session.headers
            payload = {
                "workspaceIds": workspace_ids
            }
            response = self.session.post(
                f"{self.base_url}/domains/{domain_id}/unassignWorkspaces",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return {"status": "success", "message": "Workspaces unassigned from domain successfully"}
        except Exception as e:
            logger.error(f"Error unassigning workspaces from domain: {e}")
            raise


    def update_domain(self, domain_id: str, display_name: Optional[str] = None, description: Optional[str] = None) -> Dict[str, Any]:
        """Update an existing domain"""
        try:
            headers = self.session.headers
            payload = {}
            
            if display_name:
                payload["displayName"] = display_name
            if description:
                payload["description"] = description
            
            response = self.session.patch(
                f"{self.base_url}/domains/{domain_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating domain: {e}")
            raise
    

    def apply_tags_to_item(self, workspace_id: str, item_id: str, tags: List[str]) -> Dict[str, Any]:
        """Apply tags to a specific item"""
        try:
            headers = self.session.headers
            payload = {
                "tags": tags
            }
            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}/applyTags",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            
            if response.status_code == 200:
                return {
                    "success": True,
                    "message": "Tags applied successfully",
                    "applied_tags": tags,
                    "workspace_id": workspace_id,
                    "item_id": item_id
                }
            
            if response.text.strip():
                return response.json()
            else:
                return {"success": True, "message": "Operation completed"}
                
        except Exception as e:
            logger.error(f"Error applying tags to item: {e}")
            logger.error(f"Response status: {getattr(response, 'status_code', 'Unknown')}")
            logger.error(f"Response text: {getattr(response, 'text', 'No response')}")
            raise


    def list_tags_in_tenant(self) -> Dict[str, Any]:
        """List tags in the tenant"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/tags",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error listing tags in tenant: {e}")
            raise


    def unapply_tags_from_item(self, workspace_id: str, item_id: str, tags: List[str]) -> Dict[str, Any]:
        """Remove tags from a specific item"""
        try:
            headers = self.session.headers
            payload = {
                "tags": tags
            }
            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}/unapplyTags",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            if response.status_code == 200:
                return {
                    "success": True,
                    "message": "Tags removed successfully",
                    "removed_tags": tags,
                    "workspace_id": workspace_id,
                    "item_id": item_id
                }
            if response.text.strip():
                return response.json()
            else:
                return {"success": True, "message": "Operation completed"}
                
        except Exception as e:
            logger.error(f"Error unapplying tags from item: {e}")
            logger.error(f"Response status: {getattr(response, 'status_code', 'Unknown')}")
            logger.error(f"Response text: {getattr(response, 'text', 'No response')}")
            raise


    def list_capacities(self) -> List[Dict[str, Any]]:
        """List all capacities"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/capacities",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing capacities: {e}")
            raise


    def add_workspace_role_assignment(self, workspace_id: str, principal_id: str, principal_type: str, role: str) -> Dict[str, Any]:
        """Add a role assignment to a workspace"""
        try:
            headers = self.session.headers
            payload = {
                "principal": {
                    "id": principal_id,
                    "type": principal_type
                },
                "role": role
            }
            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/roleAssignments",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error adding workspace role assignment: {e}")
            raise


    def assign_workspace_to_capacity(self, workspace_id: str, capacity_id: str) -> Dict[str, Any]:
        """Assign a workspace to a capacity"""
        try:
            headers = self.session.headers
            payload = {
                "capacityId": capacity_id
            }
            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/assignToCapacity",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            if response.status_code == 200:
                return {
                    "success": True,
                    "message": "Workspace assigned successfully",
                    "capacity_id": capacity_id,
                    "workspace_id": workspace_id
                }
            
            if response.text.strip():
                return response.json()
            else:
                return {"success": True, "message": "Operation completed"}
        except Exception as e:
            logger.error(f"Error assigning workspace to capacity: {e}")
            raise


    def create_workspace(self, display_name: str, description: Optional[str] = None, capacity_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new workspace"""
        try:
            headers = self.session.headers
            payload = {
                "displayName": display_name,
                "description": description
            }
            if capacity_id:
                payload["capacityId"] = capacity_id
            response = self.session.post(
                f"{self.base_url}/workspaces",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error creating workspace: {e}")
            raise


    def delete_workspace(self, workspace_id: str) -> None:
        """Delete a workspace"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/workspaces/{workspace_id}",
                headers=headers
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error deleting workspace: {e}")
            raise


    def delete_workspace_role_assignment(self, workspace_id: str, role_assignment_id: str) -> None:
        """Delete a role assignment from a workspace"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/workspaces/{workspace_id}/roleAssignments/{role_assignment_id}",
                headers=headers
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error deleting workspace role assignment: {e}")
            raise


    def get_workspace_by_id(self, workspace_id: str) -> Dict[str, Any]:
        """Get a workspace by its ID"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting workspace by ID: {e}")
            raise


    def get_workspace_role_assignment(self, workspace_id: str, role_assignment_id: str) -> Dict[str, Any]:
        """Get a role assignment from a workspace"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}/roleAssignments/{role_assignment_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting workspace role assignment: {e}")
            raise


    def list_workspace_role_assignments(self, workspace_id: str) -> List[Dict[str, Any]]:
        """List all role assignments in a specific workspace"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}/roleAssignments",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing workspace role assignments: {e}")
            raise


    def list_workspaces(self) -> List[Dict[str, Any]]:
        """List all workspaces"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing workspaces: {e}")
            raise


    def unassign_workspace_from_capacity(self, workspace_id: str) -> None:
        """Unassign a workspace from capacity"""
        try:
            headers = self.session.headers
            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/unassignFromCapacity",
                headers=headers
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error unassigning workspace from capacity: {e}")
            raise


    def update_workspace(self, workspace_id: str, display_name: Optional[str] = None, description: Optional[str] = None) -> Dict[str, Any]:
        """Update a workspace"""
        try:
            headers = self.session.headers
            payload = {}
            if display_name:
                payload["displayName"] = display_name
            if description:
                payload["description"] = description
            
            response = self.session.patch(
                f"{self.base_url}/workspaces/{workspace_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating workspace: {e}")
            raise


    def update_workspace_role_assignment(self, workspace_id: str, role_assignment_id: str, role: Optional[str] = None) -> Dict[str, Any]:
        """Update a workspace role assignment"""
        try:
            headers = self.session.headers
            payload = {}
            if role:
                payload["role"] = role

            response = self.session.patch(
                f"{self.base_url}/workspaces/{workspace_id}/roleAssignments/{role_assignment_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error updating workspace role assignment: {e}")
            raise


    def create_folder(self, workspace_id: str, folder_name: str, parent_folder_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new folder in a specific workspace"""
        try:
            headers = self.session.headers
            payload = {
                "displayName": folder_name
            }
            if parent_folder_id:
                payload["parentFolderId"] = parent_folder_id

            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/folders",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error creating folder in workspace: {e}")
            raise


    def delete_folder(self, workspace_id: str, folder_id: str) -> None:
        """Delete a folder in a specific workspace"""
        try:
            headers = self.session.headers
            response = self.session.delete(
                f"{self.base_url}/workspaces/{workspace_id}/folders/{folder_id}",
                headers=headers
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error deleting folder in workspace: {e}")
            raise


    def get_folder(self, workspace_id: str, folder_id: str) -> Dict[str, Any]:
        """Get details of a specific folder in a workspace"""
        try:
            headers = self.session.headers
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}/folders/{folder_id}",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error getting folder in workspace: {e}")
            raise


    def list_folders(self, workspace_id: str, root_folder_id: Optional[str] = None, recursive: Optional[bool] = False) -> List[Dict[str, Any]]:
        """List all folders in a specific workspace"""
        try:
            headers = self.session.headers
            params = {}
            if root_folder_id:
                params["rootFolderId"] = root_folder_id
            if recursive:
                params["recursive"] = str(recursive).lower()
            response = self.session.get(
                f"{self.base_url}/workspaces/{workspace_id}/folders",
                headers=headers,
                params=params
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing folders in workspace: {e}")
            raise


    def move_folder(self, workspace_id: str, folder_id: str, target_folder_id: str) -> None:
        """Move a folder to a new parent folder in a specific workspace"""
        try:
            headers = self.session.headers
            payload = {
                "targetFolderId": target_folder_id
            }

            response = self.session.post(
                f"{self.base_url}/workspaces/{workspace_id}/folders/{folder_id}/move",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error moving folder in workspace: {e}")
            raise


    def update_folder(self, workspace_id: str, folder_id: str, new_folder_name: str) -> None:
        """Update the name of a folder in a specific workspace"""
        try:
            headers = self.session.headers
            payload = {
                "displayName": new_folder_name
            }
            response = self.session.patch(
                f"{self.base_url}/workspaces/{workspace_id}/folders/{folder_id}",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
        except Exception as e:
            logger.error(f"Error updating folder in workspace: {e}")
            raise
    
    def get_data_from_lakehouse(self, sql_query: str):
        """
        Fetches data from the lakehouse table using SQL query with access token.
        Returns a list of dictionaries containing the data.
        """
        cnxn = None
        cursor = None
        tenant_id = os.getenv("TENANT_ID")
        client_id = os.getenv("CLIENT_ID")
        client_secret = os.getenv("CLIENT_SECRET")
        server = '2kg5tzezshsufoul3i7hmpw6fy-gsg2y2txhyyulbaxhw7yduba2a.datawarehouse.fabric.microsoft.com'
        database = 'FabricAdminAgent'

        required_vars = {
            "TENANT_ID": tenant_id,
            "CLIENT_ID": client_id,
            "CLIENT_SECRET": client_secret,
            "FABRIC_SQL_SERVER": server,
            "FABRIC_SQL_DATABASE": database
        }
        if not all(required_vars.values()):
            missing = [key for key, value in required_vars.items() if not value]
            print(f"Error: Missing required environment variables: {', '.join(missing)}")
            return None

        credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret
        )

        access_token = None

        try:
            token_details = credential.get_token("https://analysis.windows.net/powerbi/api/.default")
            access_token = token_details.token
        except Exception as e:
            print(f"Info: Could not get token with Power BI scope. Trying fallback. Details: {e}")
            try:
                token_details = credential.get_token("https://database.windows.net/.default")
                access_token = token_details.token
            except Exception as e2:
                print(f"Error: Failed to acquire token with both primary and fallback scopes. Details: {e2}")
                return None
            
        token_bytes = access_token.encode("utf-16-le")
        token_struct = struct.pack(f"<I{len(token_bytes)}s", len(token_bytes), token_bytes)

        connection_string = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Encrypt=yes;"
            f"TrustServerCertificate=no;"
            f"Connection Timeout=30;"
        )

        conn = pyodbc.connect(connection_string, attrs_before={1256: token_struct})
        try:
            cnxn = conn
            cursor = cnxn.cursor()
    
            print(f"Executing query: {sql_query}")
    
            cursor.execute(sql_query)
            columns = [column[0] for column in cursor.description]
            data_list = []
            rows = cursor.fetchall()
            
            for row in rows:
                row_dict = {columns[i]: value for i, value in enumerate(row)}
                data_list.append(row_dict)
                
            return {"value": data_list}
            
        except pyodbc.Error as e:
            return {"error": f"Database error: {str(e)}"}
        except Exception as ex:
            return {"error": f"Unexpected error: {str(ex)}"}
        finally:
            if cursor:
                cursor.close()
            if cnxn:
                cnxn.close()

mcp = FastMCP("Microsoft Fabric MCP")
msal_client = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)


def store_user_session(session_id: str, access_token: str, serialized_token_cache: str):
    """Stores user session data in Redis with an automatic expiration."""
    session_key = f"session:{session_id}"
    session_data = json.dumps({
        "access_token": access_token,
        "token_cache": serialized_token_cache
    })
    redis_client.setex(session_key, SESSION_TTL_SECONDS, session_data)


def get_user_session(session_id: str) -> Optional[Dict[str, Any]]:
    """Retrieves user session data from Redis."""
    session_key = f"session:{session_id}"
    session_data_str = redis_client.get(session_key)
    return json.loads(session_data_str) if session_data_str else None


def generate_and_store_api_key(session_id: str) -> str:
    """Creates a new API key, maps it to the session_id in Redis, and sets an expiration."""
    api_key = secrets.token_urlsafe(32)
    api_key_redis_key = f"apikey:{api_key}"
    redis_client.setex(api_key_redis_key, SESSION_TTL_SECONDS, session_id)
    return api_key


def get_session_from_api_key(api_key: str) -> Optional[str]:
    """Looks up a session_id from an API key in Redis."""
    api_key_redis_key = f"apikey:{api_key}"
    return redis_client.get(api_key_redis_key)



class RedirectException(Exception):
    def __init__(self, url: str):
        self.url = url


@mcp.tool()
def create_item(workspace_id: str, display_name: str, item_type: str, description: str = None, folder_id: str = None) -> str:
    """Create a new item in the specified workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        if fabric_client is None:
            return "Error: Not authenticated. Please authenticate first."
        
        response = fabric_client.create_item(workspace_id, display_name, item_type, description, folder_id)
        return f"Item created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating item: {str(e)}"


@mcp.tool()
def delete_item(workspace_id: str, item_id: str) -> str:
    """Delete a specific item from a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.delete_item(workspace_id, item_id)
        return f"Item deleted successfully"
    except Exception as e:
        return f"Error deleting item: {str(e)}"


@mcp.tool()
def get_item_definition(workspace_id: str, item_id: str, format: str = None) -> str:
    """Get the definition of a specific item"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.get_item_definition(workspace_id, item_id, format)
        return f"Item definition retrieved: {response}"
    except Exception as e:
        return f"Error getting item definition: {str(e)}"


@mcp.tool()
def list_item_connections(workspace_id: str, item_id: str) -> str:
    """List all connections for a specific item"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_item_connections(workspace_id, item_id)
        return f"Item connections: {response}"
    except Exception as e:
        return f"Error listing item connections: {str(e)}"


@mcp.tool()
def list_items(workspace_id: str, item_type: str = None, recursive: bool = False, root_folder_id: str = None) -> str:
    """List all items of a specific type in a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.list_items(workspace_id, item_type, recursive, root_folder_id)
        return f"Items found: {len(response)} items - {response}"
    except Exception as e:
        return f"Error listing items: {str(e)}"


@mcp.tool()
def update_item(workspace_id: str, item_id: str, updated_display_name: str, updated_description: str = None) -> str:
    """Update a specific item in a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.update_item(workspace_id, item_id, updated_display_name, updated_description)
        return f"Item updated successfully: {response}"
    except Exception as e:
        return f"Error updating item: {str(e)}"


@mcp.tool()
def add_role_assignment_deployment_pipeline(pipeline_id: str, role: str, principal_id: str, principal_type: str) -> str:
    """
    Add a role assignment to a deployment pipeline
    
    Args:
        pipeline_id: The ID of the deployment pipeline
        role: The role to assign (e.g., "Contributor", "Reader")
        principal_id: The ID of the user or service principal to assign the role to
        principal_type: The type of the principal (e.g., "User", "ServicePrincipal")
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        result = fabric_client.add_role_assignment_deployment_pipeline(
            pipeline_id=pipeline_id,
            role=role,
            principal_id=principal_id,
            principal_type=principal_type
        )
        
        return f"✅ Role assignment added successfully!\n" \
               f"Pipeline ID: {pipeline_id}\n" \
               f"Role: {role}\n" \
               f"Principal ID: {principal_id}\n" \
               f"Principal Type: {principal_type}"
    
    except Exception as e:
        return f"❌ Error adding role assignment: {str(e)}"


@mcp.tool()
def assign_workspace_to_deployment_pipeline_stage(pipeline_id: str, stage_id: str, workspace_id: str) -> str:
    """
    Assign a workspace to a specific stage in a deployment pipeline
   
    Args:
        pipeline_id: The ID of the deployment pipeline
        stage_id: The ID of the stage to assign the workspace to
        workspace_id: The ID of the workspace to assign
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        result = fabric_client.assign_workspace_to_a_stage_in_deployment_pipeline(
            pipeline_id=pipeline_id,
            stage_id=stage_id,
            workspace_id=workspace_id
        )
        
        if isinstance(result, dict) and result.get("success"):
            return f"✅ Workspace assigned to stage successfully!\n" \
                   f"Pipeline ID: {pipeline_id}\n" \
                   f"Stage ID: {stage_id}\n" \
                   f"Workspace ID: {workspace_id}"
        else:
            return f"✅ Workspace assigned to stage successfully!\n" \
                   f"Pipeline ID: {pipeline_id}\n" \
                   f"Stage ID: {stage_id}\n" \
                   f"Workspace ID: {workspace_id}\n" \
                   f"Response: {result}"
   
    except Exception as e:
        return f"❌ Error assigning workspace to stage: {str(e)}"


@mcp.tool()
def create_deployment_pipeline(pipeline_name: str, stage_names: List[str], description: Optional[str] = None) -> str:
    """
    Creates a deployment pipeline with a specific list of stage names.
    
    Args:
        pipeline_name: The name for the new deployment pipeline.
        stage_names: A list of names for the stages to be created, like ["Development", "Test", "Production"].
        description: Optional description for the pipeline.
    """
    try:
        if not stage_names:
            return "❌ Error: You must provide at least one stage name."

        stages = [{"order": i, "displayName": name} for i, name in enumerate(stage_names)]
        
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."

        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        
        result = fabric_client.create_deployment_pipeline(
            display_name=pipeline_name,
            description=description,
            stages=stages
        )
        
        response_text = f"✅ Deployment pipeline '{pipeline_name}' created successfully!\n"
        response_text += f"ID: {result['id']}\n"
        response_text += f"Stages ({len(stages)}):\n"
        
        for stage in stages:
            response_text += f"  - {stage['displayName']}\n"
        
        if description:
            response_text += f"Description: {description}\n"

        response_text += f"\nURL: https://app.fabric.microsoft.com/pipelines/{result['id']}?experience=fabric-developer"

        return response_text
    
    except Exception as e:
        return f"❌ Error creating deployment pipeline: {str(e)}"


@mcp.tool()
def delete_deployment_pipeline(pipeline_id: str) -> str:
    """Delete a specific deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.delete_deployment_pipeline(pipeline_id)
        return f"Deployment pipeline deleted successfully"
    except Exception as e:
        return f"Error deleting deployment pipeline: {str(e)}"


@mcp.tool()
def delete_deployment_pipeline_role_assignment(pipeline_id: str, principal_id: str) -> str:
    """Delete a role assignment from a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.delete_deployment_pipeline_role_assignment(pipeline_id, principal_id)
        return f"Role Assignment deleted successfully"
    except Exception as e:
        return f"Error deleting deployment pipeline role assignment: {str(e)}"


@mcp.tool()
def deploy_stage_content(pipeline_id: str, source_stage_id: str, target_stage_id: str, note: str = None, items: List[Dict] = None) -> str:
    """Deploy content to a specific stage in a deployment pipeline
    
    Args:
        pipeline_id: The deployment pipeline ID
        source_stage_id: The ID of the source stage
        target_stage_id: The ID of the target stage  
        note: Optional note describing the deployment (max 1024 characters)
        items: Optional list of items to deploy. Each item should have:
               - sourceItemId: The ID of the item to deploy
               - itemType: The type of the item (e.g., 'Report', 'Dashboard', 'SemanticModel')
               If not provided, all supported items will be deployed
    
    Returns:
        Success message with deployment details
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            return "Error: Session not found or expired. Please re-authenticate."

        fabric_client = FabricClient(user_session['access_token'])

        # Validate input parameters
        if not pipeline_id or not source_stage_id or not target_stage_id:
            return "Error: pipeline_id, source_stage_id, and target_stage_id are required"
        
        # Validate note length
        if note and len(note) > 1024:
            return "Error: Note cannot exceed 1024 characters"
        
        # Validate items structure if provided
        if items:
            for i, item in enumerate(items):
                if not item.get('sourceItemId') and not item.get('id'):
                    return f"Error: Item {i} is missing sourceItemId or id"
                if not item.get('itemType') and not item.get('type'):
                    return f"Error: Item {i} is missing itemType or type"

        response = fabric_client.deploy_stage_content(pipeline_id, source_stage_id, target_stage_id, note, items)
        
        if response.get('status_code') == 202:
            return f"Deployment initiated successfully. Status: {response.get('status', 'In Progress')}"
        else:
            return f"Stage content deployed successfully: {response}"
            
    except Exception as e:
        return f"Error deploying stage content: {str(e)}"


@mcp.tool()
def get_deployment_pipeline(pipeline_id: str) -> str:
    """Get details of a specific deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.get_deployment_pipeline(pipeline_id)
        return f"Deployment pipeline details: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline details: {str(e)}"


@mcp.tool()
def get_deployment_pipeline_stage(pipeline_id: str, stage_id: str) -> str:
    """Get details of a specific stage in a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.get_deployment_pipeline_stage(pipeline_id, stage_id)
        return f"Deployment pipeline stage details: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline stage details: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_operations(pipeline_id: str) -> str:
    """List all operations for a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_deployment_pipeline_operations(pipeline_id)
        return f"Deployment pipeline operations: {response}"
    except Exception as e:
        return f"Error listing deployment pipeline operations: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_role_assignments(pipeline_id: str) -> str:
    """List all role assignments for a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_deployment_pipeline_role_assignments(pipeline_id)
        return f"Deployment pipeline role assignments: {response}"
    except Exception as e:
        return f"Error listing role assignments: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_stage_items(pipeline_id: str, stage_id: str) -> str:
    """List all items in a specific stage of a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.list_deployment_pipeline_stage_items(pipeline_id, stage_id)
        return f"Deployment pipeline stage items: {response}"
    except Exception as e:
        return f"Error listing deployment pipeline stage items: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_stages(pipeline_id: str) -> str:
    """List stages of a specific deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.list_deployment_pipeline_stages(pipeline_id)
        return f"Deployment pipeline stages: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline stages: {str(e)}"


@mcp.tool()
def list_deployment_pipelines() -> str:
    """List all deployment pipelines"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_deployment_pipelines()
        return f"Deployment pipelines: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing deployment pipelines: {str(e)}"


@mcp.tool()
def unassign_workspace_from_deployment_pipeline_stage(pipeline_id: str, stage_id: str) -> str:
    """Unassign a workspace from a specific stage in a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.unassign_workspace_from_a_stage_in_deployment_pipeline(pipeline_id, stage_id)
        return f"Workspace unassigned from pipeline stage successfully: {response}"
    except Exception as e:
        return f"Error unassigning workspace from deployment pipeline: {str(e)}"


@mcp.tool()
def update_deployment_pipeline(pipeline_id: str, display_name: Optional[str] = None, description: Optional[str] = None, stages: Optional[List[Dict[str, Any]]] = None) -> str:
    """
    Update an existing deployment pipeline
    
    Args:
        pipeline_id: The ID of the deployment pipeline to update
        display_name: New display name for the pipeline
        description: New description for the pipeline
        stages: Updated list of stages for the pipeline
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        result = fabric_client.update_deployment_pipeline(
            pipeline_id=pipeline_id,
            display_name=display_name,
            description=description,
            stages=stages
        )
        
        response_text = f"✅ Deployment pipeline '{result['displayName']}' updated successfully!\n"
        response_text += f"ID: {result['id']}\n"
        
        if display_name:
            response_text += f"New Display Name: {display_name}\n"
        
        if description:
            response_text += f"New Description: {description}\n"
        
        if stages is not None:
            response_text += f"Stages ({len(stages)}):\n"
            for i, stage in enumerate(stages):
                response_text += f"  {i+1}. {stage['displayName']}\n"
        
        return response_text
    
    except Exception as e:
        return f"❌ Error updating deployment pipeline: {str(e)}"


@mcp.tool()
def update_deployment_pipeline_stage(pipeline_id: str, stage_id: str, display_name: str = None, description: str = None) -> str:
    """Update a specific stage in a deployment pipeline"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.update_deployment_pipeline_stage(pipeline_id, stage_id, display_name, description)
        return f"Deployment pipeline stage updated successfully: {response}"
    except Exception as e:
        return f"Error updating deployment pipeline stage: {str(e)}"


@mcp.tool()
def assign_workspace_to_domain_by_capacity(domain_id: str, capacity_id: str) -> str:
    """Assign a workspace to a domain by capacity"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.assign_workspace_to_domain_by_capacity(domain_id, capacity_id)
        return f"Workspace assigned to domain by capacity successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to domain: {str(e)}"


@mcp.tool()
def assign_workspace_to_domain_by_ids(domain_id: str, workspace_ids: List[str]) -> str:
    """Assign workspaces to a domain by IDs"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.assign_workspace_to_domain_by_ids(domain_id, workspace_ids)
        return f"Workspaces assigned to domain successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to domain: {str(e)}"


@mcp.tool()
def create_domain(domain_name: str, description: str = None, parent_id: str = None) -> str:
    """Create a new domain in Microsoft Fabric"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.create_domain(domain_name, description, parent_id)
        return f"Domain created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating domain: {str(e)}"


@mcp.tool()
def delete_domain(domain_id: str) -> str:
    """Delete a domain in Microsoft Fabric"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.delete_domain(domain_id)
        return f"Domain deleted successfully"
    except Exception as e:
        return f"Error deleting domain: {str(e)}"


@mcp.tool()
def get_domain(domain_id: str) -> str:
    """Get details of a specific domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.get_domain(domain_id)
        return f"Domain details: {response}"
    except Exception as e:
        return f"Error getting domain details: {str(e)}"


@mcp.tool()
def list_domain_workspaces(domain_id: str) -> str:
    """List all workspaces in a specific domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_domain_workspaces(domain_id)
        return f"Domain workspaces: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing domain workspaces: {str(e)}"


@mcp.tool()
def list_domains() -> str:
    """List all domains in Microsoft Fabric"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_domains()
        return f"Domains: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing domains: {str(e)}"


@mcp.tool()
def domain_bulk_assign_roles(domain_id: str, role: str, principals: List[Dict]) -> str:
    """Bulk assign roles to principals in a domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.domain_bulk_assign_roles(domain_id, role, principals)
        return f"Roles bulk assigned successfully: {response}"
    except Exception as e:
        return f"Error bulk assigning roles in domain: {str(e)}"


@mcp.tool()
def domain_bulk_unassign_roles(domain_id: str, role: str, principals: List[Dict]) -> str:
    """Bulk unassign roles from principals in a domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.domain_bulk_unassign_roles(domain_id, role, principals)
        return f"Roles bulk unassigned successfully: {response}"
    except Exception as e:
        return f"Error bulk unassigning roles in domain: {str(e)}"


@mcp.tool()
def domain_unassign_all_workspaces(domain_id: str) -> str:
    """Unassign all workspaces from a domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.domain_unassign_all_workspaces(domain_id)
        return f"All workspaces unassigned from domain successfully: {response}"
    except Exception as e:
        return f"Error unassigning all workspaces from domain: {str(e)}"


@mcp.tool()
def domain_unassign_workspace_by_ids(domain_id: str, workspace_ids: List[str]) -> str:
    """Unassign specific workspaces from a domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.domain_unassign_workspace_by_ids(domain_id, workspace_ids)
        return f"Workspaces unassigned from domain successfully: {response}"
    except Exception as e:
        return f"Error unassigning workspaces from domain: {str(e)}"


@mcp.tool()
def update_domain(domain_id: str, display_name: str = None, description: str = None) -> str:
    """Update an existing domain"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.update_domain(domain_id, display_name, description)
        return f"Domain updated successfully: {response}"
    except Exception as e:
        return f"Error updating domain: {str(e)}"


@mcp.tool()
def apply_tags_to_item(workspace_id: str, item_id: str, tags_ids: List[str]) -> str:
    """Apply tags to a specific item"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        result = fabric_client.apply_tags_to_item(workspace_id=workspace_id, item_id=item_id, tags=tags_ids)
        return f"✅ Tags applied to item '{item_id}' successfully!\nTags: {', '.join(tags_ids)}"
    except Exception as e:
        return f"❌ Error applying tags to item: {str(e)}"


@mcp.tool()
def list_tags_in_tenant() -> str:
    """List tags in the tenant"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_tags_in_tenant()
        return f"Tenant tags: {response}"
    except Exception as e:
        return f"Error listing tags in tenant: {str(e)}"


@mcp.tool()
def unapply_tags_from_item(workspace_id: str, item_id: str, tags_ids: List[str]) -> str:
    """Remove tags from a specific item"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.unapply_tags_from_item(workspace_id, item_id, tags_ids)
        return f"Tags removed successfully: {response}"
    except Exception as e:
        return f"Error unapplying tags from item: {str(e)}"


@mcp.tool()
def list_capacities() -> str:
    """List all capacities"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.list_capacities()
        return f"Capacities: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing capacities: {str(e)}"


@mcp.tool()
def add_workspace_role_assignment(workspace_id: str, principal_id: str, principal_type: str, role: str) -> str:
    """Add a role assignment to a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.add_workspace_role_assignment(workspace_id, principal_id, principal_type, role)
        return f"Workspace role assignment added successfully: {response}"
    except Exception as e:
        return f"Error adding workspace role assignment: {str(e)}"


@mcp.tool()
def assign_workspace_to_capacity(workspace_id: str, capacity_id: str) -> str:
    """Assign a workspace to a capacity"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.assign_workspace_to_capacity(workspace_id, capacity_id)
        return f"Workspace assigned to capacity successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to capacity: {str(e)}"


@mcp.tool()
def create_workspace(display_name: str, description: str = None, capacity_id: str = None) -> str:
    """Create a new workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.create_workspace(display_name, description, capacity_id)
        return f"Workspace created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating workspace: {str(e)}"


@mcp.tool()
def delete_workspace(workspace_id: str) -> str:
    """Delete a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        fabric_client.delete_workspace(workspace_id)
        return f"Workspace deleted successfully"
    except Exception as e:
        return f"Error deleting workspace: {str(e)}"


@mcp.tool()
def delete_workspace_role_assignment(workspace_id: str, role_assignment_id: str) -> str:
    """Delete a role assignment from a workspace"""
    try:
        session_id = current_user_context.get()
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        fabric_client.delete_workspace_role_assignment(workspace_id, role_assignment_id)
        return f"Workspace role assignment deleted successfully"
    except Exception as e:
        return f"Error deleting workspace role assignment: {str(e)}"


@mcp.tool()
def get_workspace_by_id(workspace_id: str) -> str:
    """Get a workspace by its ID"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.get_workspace_by_id(workspace_id)
        return f"Workspace details: {response}"
    except Exception as e:
        return f"Error getting workspace by ID: {str(e)}"


@mcp.tool()
def get_workspace_role_assignment(workspace_id: str, role_assignment_id: str) -> str:
    """Get a role assignment from a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.get_workspace_role_assignment(workspace_id, role_assignment_id)
        return f"Workspace role assignment: {response}"
    except Exception as e:
        return f"Error getting workspace role assignment: {str(e)}"


@mcp.tool()
def list_workspace_role_assignments(workspace_id: str) -> str:
    """List all role assignments in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_workspace_role_assignments(workspace_id)
        return f"Workspace role assignments: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing workspace role assignments: {str(e)}"


@mcp.tool()
def list_workspaces() -> str:
    """List all workspaces"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])
        response = fabric_client.list_workspaces()
        return f"Workspaces: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing workspaces: {str(e)}"


@mcp.tool()
def unassign_workspace_from_capacity(workspace_id: str) -> str:
    """Unassign a workspace from capacity"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        fabric_client.unassign_workspace_from_capacity(workspace_id)
        return f"Workspace unassigned from capacity successfully"
    except Exception as e:
        return f"Error unassigning workspace from capacity: {str(e)}"


@mcp.tool()
def update_workspace(workspace_id: str, display_name: str = None, description: str = None) -> str:
    """Update a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.update_workspace(workspace_id, display_name, description)
        return f"Workspace updated successfully: {response}"
    except Exception as e:
        return f"Error updating workspace: {str(e)}"


@mcp.tool()
def update_workspace_role_assignment(workspace_id: str, role_assignment_id: str, role: str = None) -> str:
    """Update a workspace role assignment"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.update_workspace_role_assignment(workspace_id, role_assignment_id, role)
        return f"Workspace role assignment updated successfully: {response}"
    except Exception as e:
        return f"Error updating workspace role assignment: {str(e)}"


@mcp.tool()
def create_folder(workspace_id: str, folder_name: str, parent_folder_id: str = None) -> str:
    """Create a new folder in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.create_folder(workspace_id, folder_name, parent_folder_id)
        return f"Folder created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating folder in workspace: {str(e)}"


@mcp.tool()
def delete_folder(workspace_id: str, folder_id: str) -> str:
    """Delete a folder in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        fabric_client.delete_folder(workspace_id, folder_id)
        return f"Folder deleted successfully"
    except Exception as e:
        return f"Error deleting folder in workspace: {str(e)}"


@mcp.tool()
def get_folder(workspace_id: str, folder_id: str) -> str:
    """Get details of a specific folder in a workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.get_folder(workspace_id, folder_id)
        return f"Folder details: {response}"
    except Exception as e:
        return f"Error getting folder in workspace: {str(e)}"


@mcp.tool()
def list_folders(workspace_id: str, root_folder_id: str = None, recursive: bool = False) -> str:
    """List all folders in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        response = fabric_client.list_folders(workspace_id, root_folder_id, recursive)
        return f"Folders: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing folders in workspace: {str(e)}"


@mcp.tool()
def move_folder(workspace_id: str, folder_id: str, target_folder_id: str) -> str:
    """Move a folder to a new parent folder in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        fabric_client.move_folder(workspace_id, folder_id, target_folder_id)
        return f"Folder moved successfully"
    except Exception as e:
        return f"Error moving folder in workspace: {str(e)}"


@mcp.tool()
def update_folder(workspace_id: str, folder_id: str, new_folder_name: str) -> str:
    """Update the name of a folder in a specific workspace"""
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."
        
        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        fabric_client.update_folder(workspace_id, folder_id, new_folder_name)
        return f"Folder updated successfully"
    except Exception as e:
        return f"Error updating folder in workspace: {str(e)}"


@mcp.tool()
def get_table_schema(table_name: str) -> str:
    """
    Retrieves the schema (column names and data types) of a specified table from the lakehouse.
    The LLM should use this tool to discover the structure of a table before generating a SQL query.

    Args:
        table_name: The name of the table to retrieve the schema for. 
                    Valid table names are: 'datasetrefreshanalysis', 'workspaces_cu_consumption', 
                    'orphan_unused_workspaces', 'items_cu_consumption', 'orphan_unused_datasets', 
                    'orphan_unused_reports', 'outlier_items'.

    Returns:
        A string describing the table schema or an error message if the table is not found.
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."

        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        query = f"SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table_name}'"
        
        result = fabric_client.get_data_from_lakehouse(query)

        if result.get("error") or not result.get("value"):
            return f"Error: Could not retrieve schema for table '{table_name}'. Please ensure the table name is correct."

        schema_info = f"Schema for table '{table_name}':\n"
        for row in result.get("value", []):
            schema_info += f"- {row['COLUMN_NAME']} ({row['DATA_TYPE']})\n"

        return schema_info

    except Exception as e:
        return f"An unexpected error occurred while retrieving the schema for table '{table_name}': {str(e)}"
    

@mcp.tool()
def query_lakehouse_and_get_insights(sql_query: str) -> str:
    """
    Executes a SQL query against the Microsoft Fabric lakehouse to retrieve insights.
    The LLM should generate the SQL query based on the user's request and the available tables.

    Available Tables and their Description:
    1.  `datasetrefreshanalysis`: Contains analysis of dataset refresh operations.
    2.  `workspaces_cu_consumption`: Tracks capacity unit (CU) consumption by workspace.
    3.  `orphan_unused_workspaces`: Identifies workspaces that are either orphaned (not assigned to any fabric capacity) or unused.
    4.  `items_cu_consumption`: Contains capacity unit (CU) consumption details by fabric artifacts.
    5.  `orphan_unused_datasets`: Identifies datasets that are either orphaned (not been used in any Report) or unused.
    6.  `orphan_unused_reports`: Identifies items that are either orphaned (do not have any semantic model attached) or unused.
    7.  `outlier_items`: Contains information about items that are considered outliers in terms of capacity unit (CU) consumption.
    
    Args:
        sql_query: The SQL query to execute on the lakehouse.

    Returns:
        A string containing the results of the query or an error message.
    """
    try:
        session_id = current_user_context.get()
        if not session_id:
            return "Error: No user context found. Please authenticate first."

        user_session = get_user_session(session_id)
        if not user_session:
            raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

        fabric_client = FabricClient(user_session['access_token'])

        result = fabric_client.get_data_from_lakehouse(sql_query)

        return f"Insights retrieved successfully: {result}"
    except Exception as e:
        return f"Error retrieving insights from lakehouse: {str(e)}"


@mcp.tool()
def check_config() -> str:
    """Check if the authentication configuration is properly set up"""
    config_status = []

    if CLIENT_ID and TENANT_ID and CLIENT_SECRET:
        config_status.append(f"✅ Environment variables (CLIENT_ID, TENANT_ID, CLIENT_SECRET) are set.")
    else:
        config_status.append("❌ Critical environment variables are missing. Please check your .env file.")
        return "\n".join(config_status)

    session_id = current_user_context.get()
    if not session_id:
        return "Error: No user context found. Please authenticate first."
    
    user_session = get_user_session(session_id)
    if not user_session:
        raise HTTPException(status_code=401, detail="Session not found or expired. Please re-authenticate.")

    fabric_client = FabricClient(user_session['access_token'])
    if fabric_client and fabric_client.access_token:
        config_status.append("✅ Signed in successfully (found a valid access token).")
    else:
        config_status.append("❌ User is not signed in or the session has expired.")
    
    status_text = "\n".join(config_status)
    if "✅" in status_text and "❌" not in status_text:
        status_text += "\n\n✅ Configuration appears complete and you are signed in!"
    else:
        status_text += "\n\n❌ Configuration is incomplete or you are not signed in."
    
    return status_text

mcp_app = mcp.http_app()

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    Ensure all services, including MCP, are initialized before the app starts serving.
    """
    global initialization_complete
    
    try:
        redis_client.ping()
        print("INFO:     Successfully connected to Redis.")
    except redis.exceptions.ConnectionError as e:
        print(f"ERROR:    Could not connect to Redis: {e}")
        raise

    async with mcp_app.lifespan(app) as mcp_lifespan_manager:
        print("INFO:     MCP App startup complete.")
        yield
        print("INFO:     MCP App shutdown starting.")
    print("INFO:     MCP App shutdown complete.")

app = FastAPI(lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET_KEY)

class RedirectException(Exception):
    def __init__(self, url: str):
        self.url = url

@app.exception_handler(RedirectException)
async def redirect_exception_handler(request: Request, exc: RedirectException):
    return RedirectResponse(url=exc.url)


@app.middleware("http")
async def mcp_auth_middleware(request: Request, call_next):
    context_token = None
    if request.url.path.startswith('/mcp-server'):
        api_key = request.headers.get("X-API-Key") or request.headers.get("Authorization", "").replace("Bearer ", "")
        if api_key:
            session_id = get_session_from_api_key(api_key)
            if session_id:
                context_token = current_user_context.set(session_id)
    
    response = await call_next(request)

    if context_token:
        current_user_context.reset(context_token)
    return response


@app.middleware("http")
async def add_safe_session_header(request: Request, call_next):
    response = await call_next(request)
    if "Mcp-Session-Id" in response.headers:
        response.headers["X-Mcp-Session-Id"] = response.headers["Mcp-Session-Id"]
    return response


@app.get("/login")
def login(request: Request):
    auth_url = msal_client.get_authorization_request_url(scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return RedirectResponse(auth_url)


@app.get(REDIRECT_PATH)
def auth_callback(request: Request, code: str):
    token_cache = msal.SerializableTokenCache()
    msal_client.token_cache = token_cache
    result = msal_client.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

    if "error" in result:
        return {"error": result.get("error_description")}

    session_id = str(uuid.uuid4())
    request.session['session_id'] = session_id

    store_user_session(session_id, result['access_token'], token_cache.serialize())
    
    return RedirectResponse(url="/")


@app.get("/")
def root(request: Request):
    session_id = request.session.get('session_id')
    if session_id and get_user_session(session_id):
        api_key = generate_and_store_api_key(session_id)
        return {
            "status": "authenticated",
            "message": "You are logged in. Use the API key below for your client.",
            "api_key": api_key,
            "expires_in_minutes": SESSION_TTL_SECONDS // 60
        }
    return {"status": "unauthenticated", "login_url": "/login"}


app.mount("/mcp-server", mcp_app)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)