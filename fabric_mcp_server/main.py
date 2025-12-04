import logging
import os
import re
import requests
import uvicorn
import json
import time
from collections import defaultdict, deque
from contextlib import asynccontextmanager
from dotenv import load_dotenv
from fastapi import FastAPI
from fastmcp import FastMCP
from typing import Any, Dict, List, Optional

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("fabric-fastmcp")

class FabricClient:
    """Client for Microsoft Fabric REST API"""
    def __init__(self):
        self.tenant_id = os.getenv("FABRIC_TENANT_ID")
        self.client_id = os.getenv("FABRIC_CLIENT_ID")
        self.client_secret = os.getenv("FABRIC_CLIENT_SECRET")
        self.access_token = None
        self.token_expires_at = None
        self.base_url = "https://api.fabric.microsoft.com/v1"
        self.session = requests.Session()
        
        if not all([self.tenant_id, self.client_id, self.client_secret]):
            missing = []
            if not self.tenant_id:
                missing.append("FABRIC_TENANT_ID")
            if not self.client_id:
                missing.append("FABRIC_CLIENT_ID")
            if not self.client_secret:
                missing.append("FABRIC_CLIENT_SECRET")
            
            logger.warning(f"Missing environment variables: {', '.join(missing)}")

    def _is_token_valid(self) -> bool:
        """Check if the current access token is valid based on expiration time"""
        if not self.access_token:
            logger.info("No access token available")
            return False
        
        if not self.token_expires_at:
            logger.info("No token expiration time available")
            return False
        
        buffer_time = 300
        current_time = time.time()
        is_valid = current_time < (self.token_expires_at - buffer_time)
        
        if not is_valid:
            logger.info(f"Token expired. Current time: {current_time}, Token expires at: {self.token_expires_at}")
        
        return is_valid

    def _validate_token_with_api(self) -> bool:
        """Validate token by making a simple API call"""
        if not self.access_token:
            logger.info("No access token to validate")
            return False
        
        try:
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            response = self.session.get(
                f"{self.base_url}/workspaces",
                headers=headers,
                timeout=5
            )
            
            if response.status_code >= 401 and response.status_code < 500 and response.status_code != 400:
                logger.info(f"Token validation failed: {response.status_code} - {response.reason}")
                self.access_token = None
                self.token_expires_at = None
                return False
            
            if response.status_code == 200:
                logger.debug("Token validation successful")
                return True
            
            logger.warning(f"Token validation returned status {response.status_code}, assuming valid")
            return True
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Token validation failed due to network error: {e}")
            return False

    def get_access_token(self) -> str:
        """Get access token using client credentials flow"""
        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise ValueError("Missing required authentication parameters")
        
        if self._is_token_valid():
            if self._validate_token_with_api():
                logger.info("Using existing valid access token")
                return self.access_token
            else:
                logger.info("Token failed API validation, getting new token")
        
        logger.info("Getting new access token")
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://api.fabric.microsoft.com/.default"
        }
        
        response = self.session.post(token_url, data=data)
        response.raise_for_status()
        
        token_data = response.json()
        self.access_token = token_data["access_token"]
        
        expires_in = token_data.get("expires_in", 3600)
        self.token_expires_at = time.time() + expires_in
        
        logger.info(f"New access token acquired, expires in {expires_in} seconds")
        return self.access_token


    def get_headers(self) -> Dict[str, str]:
        """Get authorization headers with automatic token refresh"""
        self.get_access_token()
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

    def _make_authenticated_request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Make an authenticated request with automatic token refresh on 401"""
        headers = self.get_headers()
        
        if 'headers' in kwargs:
            kwargs['headers'].update(headers)
        else:
            kwargs['headers'] = headers
        
        response = self.session.request(method, url, **kwargs)
        
        if response.status_code == 401:
            logger.warning("Received 401 Unauthorized, refreshing token and retrying...")
            self.access_token = None
            self.token_expires_at = None
            headers = self.get_headers()
            kwargs['headers'].update(headers)
            response = self.session.request(method, url, **kwargs)
            
            if response.status_code == 401:
                logger.error("Still receiving 401 after token refresh - check credentials")
        
        return response

    def refresh_token(self) -> str:
        """Force refresh the access token"""
        logger.info("Forcing token refresh")
        self.access_token = None
        self.token_expires_at = None
        return self.get_access_token()

    def create_item(self, workspace_id: str, displayName: str, item_type: str, description: Optional[str] = None, folder_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new item in the specified workspace"""
        try:
            headers = self.get_headers()
            
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
            headers = self.get_headers()
            response = self.session.delete(
                f"{self.base_url}/workspaces/{workspace_id}/items/{item_id}",
                headers=headers
            )
            response.raise_for_status()
            
            # Handle the response based on status code and content
            if response.status_code == 200:
                # API returns empty JSON object ({}) on successful deletion
                try:
                    json_response = response.json()
                    # Return success message with the empty response
                    return {
                        "success": True,
                        "message": "Item deleted successfully",
                        "workspace_id": workspace_id,
                        "item_id": item_id,
                        "response": json_response
                    }
                except ValueError:
                    # In case response is not valid JSON
                    return {
                        "success": True,
                        "message": "Item deleted successfully",
                        "workspace_id": workspace_id,
                        "item_id": item_id
                    }
            else:
                # For other successful status codes
                return {
                    "success": True,
                    "message": "Item deleted successfully",
                    "workspace_id": workspace_id,
                    "item_id": item_id,
                    "status_code": response.status_code
                }
                
        except Exception as e:
            logger.error(f"Error deleting item: {e}")
            raise


    def get_item_definition(self, workspace_id: str, item_id: str, type: str) -> Dict[str, Any]:
        """Get the definition of a specific item"""
        try:
            headers = self.get_headers()
            url = f"{self.base_url}/admin/workspaces/{workspace_id}/items/{item_id}"
            if type:
                url += f"?type={type}"
            response = self.session.get(
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
            
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
            headers = self.get_headers()
            
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
            headers = self.get_headers()
            
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
            headers = self.get_headers()
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
            headers = self.get_headers()
                        
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
            headers = self.get_headers()
            payload = {
                "sourceStageId": source_stage_id,
                "targetStageId": target_stage_id,
                "note": note if note else ""
            }
            
            if items:
                payload["items"] = [
                    {
                        "sourceItemId": item['id'],
                        "itemType": item['type']
                    }
                    for item in items
                ] 
                
            response = self.session.post(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/deploy",
                headers=headers,
                json=payload
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Error deploying stage content: {e}")
            raise


    def get_deployment_pipeline(self, pipeline_id: str) -> Dict[str, Any]:
        """Get details of a specific deployment pipeline"""
        try:
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}/items",
                headers=headers
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except Exception as e:
            logger.error(f"Error listing deployment pipeline stage items: {e}")
            raise


    def list_deployment_pipeline_stages(self, pipeline_id: str, stage_id: str) -> List[Dict[str, Any]]:
        """List stages of a specific deployment pipeline"""
        try:
            headers = self.get_headers()
            response = self.session.get(
                f"{self.base_url}/deploymentPipelines/{pipeline_id}/stages/{stage_id}",
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
            headers = self.get_headers()
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
            headers = self.get_headers()
            
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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
            headers = self.get_headers()
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

    # Add these methods to your FabricClient class

    def _normalize_name(self, name: str) -> str:
        """Normalize dataset/semantic model name for comparison"""
        if not isinstance(name, str):
            return ""
        return name.strip().lower()

    def _is_guid(self, s: str) -> bool:
        """Check if string is a valid GUID"""
        if not isinstance(s, str):
            return False
        return bool(re.match(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$", s))

    def get_powerbi_datasets(self, workspace_id: str) -> List[Dict[str, Any]]:
        """Get all Power BI datasets in a workspace"""
        try:
            logger.info(f"Fetching datasets from workspace: {workspace_id}")
            datasets = []
            datasets_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
            
            headers = self.get_headers()
            response = self.session.get(datasets_url, headers=headers, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            datasets = data.get("value", [])
            logger.info(f"Retrieved {len(datasets)} datasets from workspace {workspace_id}")
            return datasets
        except Exception as e:
            logger.error(f"Error getting Power BI datasets: {e}")
            raise

    def get_dataset_datasources(self, workspace_id: str, dataset_id: str) -> List[Dict[str, Any]]:
        """Get datasources for a specific dataset with timeout and error handling"""
        try:
            headers = self.get_headers()
            url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/datasources"
            
            # Add timeout and better error handling
            response = self.session.get(url, headers=headers, timeout=10)
            
            if response.status_code == 404:
                logger.debug(f"Dataset {dataset_id} not found or no datasources")
                return []
            elif response.status_code == 403:
                logger.debug(f"Access denied for dataset {dataset_id} datasources")
                return []
            elif response.status_code != 200:
                logger.warning(f"Could not get datasources for dataset {dataset_id}: {response.status_code}")
                return []
                
            return response.json().get("value", [])
        except Exception as e:
            logger.debug(f"Error getting datasources for dataset {dataset_id}: {e}")
            return []

    def get_powerbi_reports(self, workspace_id: str) -> List[Dict[str, Any]]:
        """Get all Power BI reports in a workspace"""
        try:
            logger.info(f"Fetching reports from workspace: {workspace_id}")
            reports = []
            reports_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports"
            
            headers = self.get_headers()
            response = self.session.get(reports_url, headers=headers, timeout=30)
            response.raise_for_status()

            data = response.json()
            reports = data.get("value", [])
            logger.info(f"Retrieved {len(reports)} reports from workspace {workspace_id}")
            
            return reports
        except Exception as e:
            logger.error(f"Error getting Power BI reports: {e}")
            raise

    def analyze_semantic_model_dependencies(self, workspace_id: str, dataset_id: str) -> Dict[str, Any]:
        """Analyze semantic model dependencies and create lineage data"""
        try:
            logger.info(f"Starting dependency analysis for dataset ID: {dataset_id}")
            
            # Validate dataset ID format
            if not self._is_guid(dataset_id):
                raise ValueError(f"Invalid dataset ID format: {dataset_id}")
            
            # Get datasets and reports
            datasets = self.get_powerbi_datasets(workspace_id)
            reports = self.get_powerbi_reports(workspace_id)
            logger.info(f"Found {len(datasets)} datasets and {len(reports)} reports")
            
            # Verify target dataset exists
            target_dataset = next((d for d in datasets if d["id"] == dataset_id), None)
            if not target_dataset:
                raise ValueError(f"Dataset with ID '{dataset_id}' not found")
            
            logger.info(f"Target dataset found: {target_dataset['name']}")
            
            # Build dataset name lookup
            dataset_lookup = {d["id"]: d["name"] for d in datasets}
            name_to_ids = defaultdict(list)
            for d in datasets:
                normalized_name = self._normalize_name(d["name"])
                if normalized_name:  # Only add non-empty names
                    name_to_ids[normalized_name].append(d["id"])
            
            logger.info(f"Built lookup tables for {len(name_to_ids)} unique dataset names")
            
            # Find dependencies with progress logging and batching
            dependencies = defaultdict(list)
            processed_count = 0
            batch_size = 25  # Process in smaller batches to avoid timeout
            
            for i in range(0, len(datasets), batch_size):
                batch = datasets[i:i + batch_size]
                logger.info(f"Processing batch {i//batch_size + 1}/{(len(datasets) + batch_size - 1)//batch_size} ({len(batch)} datasets)")
                
                for dataset in batch:
                    try:
                        datasources = self.get_dataset_datasources(workspace_id, dataset["id"])
                        processed_count += 1
                        
                        if processed_count % 50 == 0:
                            logger.info(f"Processed {processed_count}/{len(datasets)} datasets")
                        
                        for source in datasources:
                            # Only process Power BI and Analysis Services connections
                            datasource_type = source.get("datasourceType", "")
                            if datasource_type not in ("AnalysisServices", "PowerBI"):
                                continue
                            
                            connection_details = source.get("connectionDetails", {})
                            db_name = connection_details.get("database")
                            
                            if db_name:
                                normalized_name = self._normalize_name(db_name)
                                if normalized_name and normalized_name in name_to_ids:
                                    for source_id in name_to_ids[normalized_name]:
                                        # Prevent self-referencing dependencies
                                        if source_id != dataset["id"]:
                                            dependencies[source_id].append(dataset["id"])
                                            logger.debug(f"Found dependency: {dataset_lookup.get(source_id, source_id)} -> {dataset['name']}")
                                            
                    except Exception as e:
                        logger.warning(f"Error processing dataset {dataset.get('name', 'Unknown')}: {e}")
                        continue
            
            logger.info(f"Completed datasource analysis. Found dependencies for {len(dependencies)} source datasets")
            
            # Find all dependents using BFS with cycle detection
            queue = deque([dataset_id])
            all_dependents = set()
            visited = set([dataset_id])
            iteration_count = 0
            max_iterations = 1000  # Safety limit
            
            logger.info("Starting dependency traversal using BFS")
            
            while queue and iteration_count < max_iterations:
                iteration_count += 1
                current = queue.popleft()
                current_name = dataset_lookup.get(current, current)
                
                # Get direct dependents of current dataset
                direct_dependents = dependencies.get(current, [])
                
                if direct_dependents:
                    logger.debug(f"Dataset '{current_name}' has {len(direct_dependents)} direct dependents")
                
                for dependent in direct_dependents:
                    if dependent not in visited:
                        visited.add(dependent)
                        all_dependents.add(dependent)
                        queue.append(dependent)
                        dependent_name = dataset_lookup.get(dependent, dependent)
                        logger.debug(f"Added dependent: {dependent_name}")
                    else:
                        logger.debug(f"Skipping already visited: {dataset_lookup.get(dependent, dependent)}")
            
            if iteration_count >= max_iterations:
                logger.warning(f"Dependency traversal stopped at {max_iterations} iterations to prevent infinite loops")
            
            logger.info(f"Dependency traversal complete after {iteration_count} iterations. Found {len(all_dependents)} dependent datasets")
            
            # Create report mapping
            dataset_to_reports = defaultdict(list)
            for report in reports:
                report_dataset_id = report.get("datasetId")
                if report_dataset_id:
                    dataset_to_reports[report_dataset_id].append(report["name"])
            
            logger.info(f"Mapped reports to datasets. Found reports for {len(dataset_to_reports)} datasets")
            
            # Build lineage data
            lineage_data = []
            all_datasets = [dataset_id] + list(all_dependents)
            
            for ds_id in all_datasets:
                dataset_name = dataset_lookup.get(ds_id, "Unknown")
                role = "Main" if ds_id == dataset_id else "Dependent"
                
                report_names = dataset_to_reports.get(ds_id, [])
                if report_names:
                    for report_name in report_names:
                        lineage_data.append({
                            "dataset_id": ds_id,
                            "dataset_name": dataset_name,
                            "report_name": report_name,
                            "role": role
                        })
                else:
                    # Include datasets without reports
                    lineage_data.append({
                        "dataset_id": ds_id,
                        "dataset_name": dataset_name,
                        "report_name": None,
                        "role": role
                    })
            
            total_reports_with_datasets = len([item for item in lineage_data if item["report_name"]])
            logger.info(f"Lineage analysis complete. Created {len(lineage_data)} lineage entries ({total_reports_with_datasets} with reports)")
            
            return {
                "success": True,
                "target_dataset_id": dataset_id,
                "target_dataset_name": target_dataset["name"],
                "workspace_id": workspace_id,
                "total_datasets_analyzed": len(datasets),
                "dependent_datasets": len(all_dependents),
                "total_reports": total_reports_with_datasets,
                "lineage_data": lineage_data
            }
            
        except Exception as e:
            logger.error(f"Error analyzing semantic model dependencies: {e}")
            raise

    def get_semantic_model_refresh_schedule(self, workspace_id: str, dataset_id: str) -> Dict[str, Any]:
        """Get the refresh schedule for a specific semantic model (dataset)"""
        try:
            headers = self.get_headers()
            url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/refreshSchedule"
            
            response = self.session.get(url, headers=headers, timeout=10)
            
            if response.status_code == 404:
                logger.debug(f"Dataset {dataset_id} not found or no refresh schedule")
                return {}
            elif response.status_code == 403:
                logger.debug(f"Access denied for dataset {dataset_id} refresh schedule")
                return {}
            elif response.status_code != 200:
                logger.warning(f"Could not get refresh schedule for dataset {dataset_id}: {response.status_code}")
                return {}
                
            return response.json()
        except Exception as e:
            logger.debug(f"Error getting refresh schedule for dataset {dataset_id}: {e}")
            return {}
        
    def takeover_semantic_model(self, workspace_id: str, dataset_id: str) -> Dict[str, Any]:
        """Take over a semantic model to enable refresh schedule updates
        
        Args:
            workspace_id: The workspace (group) ID
            dataset_id: The dataset ID
            
        Returns:
            Dict containing the API response or error details
        """
        try:
            headers = self.get_headers()
            url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/Default.TakeOver"
            
            response = self.session.post(url, headers=headers, timeout=10)
            
            # The API returns 200 with no content when successful
            if response.status_code == 200:
                return {"status": "success", "message": "Successfully took over semantic model"}
            elif response.status_code == 404:
                return {"status": "error", "message": "Dataset not found"}
            elif response.status_code == 403:
                return {"status": "error", "message": "Access forbidden. Check permissions"}
            else:
                return {"status": "error", "message": f"Failed to take over semantic model. Status code: {response.status_code}"}
                
        except Exception as e:
            error_msg = f"Error taking over semantic model {dataset_id}: {str(e)}"
            logger.error(error_msg)
            raise

    def update_semantic_model_refresh_schedule(self, workspace_id: str, dataset_id: str, refresh_schedule: Dict[str, Any]) -> Dict[str, Any]:
        """Update the refresh schedule for a specific semantic model (dataset)
        
        Args:
            workspace_id: The workspace (group) ID
            dataset_id: The dataset ID
            refresh_schedule: A dictionary containing the refresh schedule configuration in format:
                {
                    "value": {
                        "days": ["Monday", "Tuesday", ...],     # Required if updating days
                        "times": ["07:00", "16:00", ...],      # Optional, 24-hour format
                        "enabled": true/false,                  # Optional
                        "localTimeZoneId": "UTC",              # Optional
                        "notifyOption": "MailOnFailure"        # Optional
                    }
                }
                
        Returns:
            Dict containing the API response
            
        Notes:
            Follows the Power BI REST API specification for updating refresh schedules.
            See: https://learn.microsoft.com/en-us/rest/api/power-bi/datasets/update-refresh-schedule
        """
        try:
            if not isinstance(refresh_schedule, dict):
                raise ValueError("Refresh schedule must be a dictionary")

            if "value" not in refresh_schedule:
                refresh_schedule = {"value": refresh_schedule}

            value = refresh_schedule["value"]
            
            valid_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            if "days" in value:
                if not value["days"] or not all(day in valid_days for day in value["days"]):
                    raise ValueError(f"Days must be a non-empty list containing valid days: {valid_days}")

            if "times" in value:
                time_pattern = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
                if not all(bool(time_pattern.match(t)) for t in value["times"]):
                    raise ValueError("Times must be in 24-hour HH:MM format (e.g., '07:00', '16:30')")

            valid_notify_options = ["NoNotification", "MailOnFailure"]
            if "notifyOption" in value and value["notifyOption"] not in valid_notify_options:
                raise ValueError(f"NotifyOption must be one of: {valid_notify_options}")

            if "enabled" in value and not value["enabled"] and len(value) > 1:
                raise ValueError("A request that disables the refresh schedule should contain no other changes")
                
            url = f"https://api.powerbi.com/v1.0/myorg/datasets/{dataset_id}/refreshSchedule"
            headers = self.get_headers()
            
            response = self.session.patch(url, headers=headers, json=refresh_schedule)
            response.raise_for_status()
            
            return {"status": "success", "message": "Successfully updated refresh schedule"}                
        except ValueError as ve:
            error_msg = f"Validation error updating refresh schedule: {str(ve)}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        except Exception as e:
            error_msg = f"Error updating refresh schedule for dataset {dataset_id}: {str(e)}"
            logger.error(error_msg)
            raise

def parse_stages_from_prompt(prompt: str) -> List[Dict[str, Any]]:
    """Parse stage names from natural language prompt"""
    stage_patterns = [
        r'\b(?:dev|development)\b',
        r'\b(?:test|testing|tst)\b', 
        r'\b(?:prod|production|prd)\b',
        r'\b(?:uat|user acceptance|staging|stage)\b',
        r'\b(?:qa|quality assurance)\b',
        r'\b(?:pre-prod|preprod|pre-production)\b'
    ]
    
    prompt_lower = prompt.lower()
    found_stages = []
    
    if 'stages' in prompt_lower:
        stage_match = re.search(r'stages?\s+([^.!?]+)', prompt_lower, re.IGNORECASE)
        if stage_match:
            stage_text = stage_match.group(1)
            stage_names = re.split(r'[,\s]+(?:and\s+)?', stage_text.strip())
            found_stages = [name.strip().title() for name in stage_names if name.strip()]
    
    if not found_stages:
        stage_map = {
            'dev': 'Development',
            'development': 'Development', 
            'test': 'Test',
            'testing': 'Test',
            'tst': 'Test',
            'prod': 'Production',
            'production': 'Production',
            'prd': 'Production',
            'uat': 'UAT',
            'user acceptance': 'UAT',
            'staging': 'Staging',
            'stage': 'Staging',
            'qa': 'QA',
            'quality assurance': 'QA',
            'pre-prod': 'Pre-Production',
            'preprod': 'Pre-Production',
            'pre-production': 'Pre-Production'
        }
        
        for key, value in stage_map.items():
            if key in prompt_lower and value not in found_stages:
                found_stages.append(value)
    
    if not found_stages:
        number_match = re.search(r'(\d+)\s+stages?', prompt_lower)
        if number_match:
            num_stages = int(number_match.group(1))
            if num_stages == 2:
                found_stages = ['Development', 'Production']
            elif num_stages == 3:
                found_stages = ['Development', 'Test', 'Production']
            elif num_stages == 4:
                found_stages = ['Development', 'Test', 'UAT', 'Production']
    
    if found_stages:
        return [{"order": i, "displayName": stage} for i, stage in enumerate(found_stages)]
    
    return [
        {"order": 0, "displayName": "Development"},
        {"order": 1, "displayName": "Test"},
        {"order": 2, "displayName": "Production"}
    ]


mcp = FastMCP("Microsoft Fabric MCP")
fabric_client = FabricClient()


@mcp.tool()
def takeover_semantic_model(workspace_id: str, dataset_id: str) -> str:
    """Take over a semantic model to enable operations like updating refresh schedule
    
    Args:
        workspace_id: The workspace (group) ID
        dataset_id: The dataset ID
    """
    try:
        # Take over the semantic model
        result = fabric_client.takeover_semantic_model(workspace_id, dataset_id)
        
        if result.get("status") == "success":
            return f"✅ {result['message']}\nWorkspace ID: {workspace_id}\nDataset ID: {dataset_id}"
        else:
            return f"❌ Failed to take over semantic model: {result.get('message')}"
            
    except Exception as e:
        return f"❌ Error: {str(e)}"

@mcp.tool()
def update_semantic_model_refresh_schedule(workspace_id: str, dataset_id: str, refresh_schedule: str) -> str:
    """Update the refresh schedule for a specific semantic model (dataset)
    
    Args:
        workspace_id: The workspace (group) ID
        dataset_id: The dataset ID
        refresh_schedule: JSON string containing the refresh schedule configuration in the format:
            {
                "value": {
                    "days": ["Monday", "Tuesday", ...],  # Required. At least one day must be specified
                    "times": ["07:00", "16:00", ...],   # Optional. Times in 24-hour HH:MM format
                    "enabled": true/false,               # Optional. Whether schedule is active
                    "localTimeZoneId": "UTC",           # Optional. TimeZoneInfo ID
                    "notifyOption": "MailOnFailure"     # Optional. MailOnFailure or NoNotification
                }
            }
            
    Notes:
        - At least one day must be specified if updating days
        - Times must be in 24-hour HH:MM format (e.g., "07:00", "16:00")
        - A request that disables the schedule should contain no other changes
        - Service principals only support NoNotification for notifyOption
        - Valid days are: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday
        - Valid notifyOptions are: NoNotification, MailOnFailure
    """
    try:
        # Parse the JSON string into a dictionary
        if isinstance(refresh_schedule, str):
            schedule_dict = json.loads(refresh_schedule)
        else:
            schedule_dict = refresh_schedule

        # Ensure the schedule has a value wrapper
        if "value" not in schedule_dict:
            schedule_dict = {"value": schedule_dict}

        # Validate the schedule format
        value = schedule_dict["value"]

        # Validate days if present
        valid_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        if "days" in value:
            if not value["days"] or not all(day in valid_days for day in value["days"]):
                return f"❌ Validation error: Days must be a non-empty list containing valid days: {valid_days}"

        # Validate times if present
        if "times" in value:
            time_pattern = "^([01][0-9]|2[0-3]):([0-5][0-9])$"
            if not all(bool(re.match(time_pattern, t)) for t in value["times"]):
                return "❌ Validation error: Times must be in 24-hour HH:MM format (e.g., '07:00', '16:30')"

        # Validate notifyOption if present
        valid_notify_options = ["NoNotification", "MailOnFailure"]
        if "notifyOption" in value and value["notifyOption"] not in valid_notify_options:
            return f"❌ Validation error: NotifyOption must be one of: {valid_notify_options}"

        # If disabling schedule, ensure no other changes
        if "enabled" in value and not value["enabled"] and len(value) > 1:
            return "❌ Validation error: A request that disables the refresh schedule should contain no other changes"

        # Make the API call
        result = fabric_client.update_semantic_model_refresh_schedule(workspace_id, dataset_id, schedule_dict)
        return f"✅ Successfully updated refresh schedule\nWorkspace ID: {workspace_id}\nDataset ID: {dataset_id}"
            
    except json.JSONDecodeError:
        return "❌ Invalid JSON format for refresh schedule. Please check the schedule format."
    except Exception as e:
        return f"❌ Error updating refresh schedule: {str(e)}"


@mcp.tool()
def get_semantic_model_refresh_schedule(workspace_id: str, dataset_id: str) -> str:
    """Get the refresh schedule for a specific semantic model (dataset)"""
    try:
        response = fabric_client.get_semantic_model_refresh_schedule(workspace_id, dataset_id)
        if response:
            return f"Refresh schedule retrieved: {response}"
        else:
            return "No refresh schedule found or access denied."
    except Exception as e:
        return f"Error getting refresh schedule: {str(e)}"


@mcp.tool()
def create_item(workspace_id: str, display_name: str, item_type: str, description: str = None, folder_id: str = None) -> str:
    """Create a new item in the specified workspace"""
    try:
        response = fabric_client.create_item(workspace_id, display_name, item_type, description, folder_id)
        return f"Item created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating item: {str(e)}"


@mcp.tool()
def delete_item(workspace_id: str, item_id: str) -> str:
    """Delete a specific item from a workspace"""
    try:
        response = fabric_client.delete_item(workspace_id, item_id)
        if response.get("success"):
            return f"✅ Item deleted successfully!\n" \
                   f"Workspace ID: {workspace_id}\n" \
                   f"Item ID: {item_id}\n" \
                   f"Status: {response.get('message', 'Completed')}"
        else:
            return f"⚠️ Item deletion completed with response: {response}"
    except Exception as e:
        return f"❌ Error deleting item: {str(e)}"


@mcp.tool()
def get_item_definition(workspace_id: str, item_id: str, type: str) -> str:
    """
    Get the definition of a specific item from Microsoft Fabric workspace.
    Args:
        workspace_id (str): The unique identifier of the workspace containing the item
        item_id (str): The unique identifier of the specific item to retrieve
        type (str): The type of the item (e.g., 'Dataset', 'Report', 'Dashboard', etc.)
    """
    try:
        response = fabric_client.get_item_definition(workspace_id, item_id, type)
        return f"Item definition retrieved: {response}"
    except Exception as e:
        return f"Error getting item definition: {str(e)}"


@mcp.tool()
def list_item_connections(workspace_id: str, item_id: str) -> str:
    """List all connections for a specific item"""
    try:
        response = fabric_client.list_item_connections(workspace_id, item_id)
        return f"Item connections: {response}"
    except Exception as e:
        return f"Error listing item connections: {str(e)}"


@mcp.tool()
def list_items(workspace_id: str, item_type: str = None, recursive: bool = False, root_folder_id: str = None) -> str:
    """List all items of a specific type in a workspace"""
    try:
        response = fabric_client.list_items(workspace_id, item_type, recursive, root_folder_id)
        return f"Items found: {len(response)} items - {response}"
    except Exception as e:
        return f"Error listing items: {str(e)}"

@mcp.tool()
def update_item(workspace_id: str, item_id: str, updated_display_name: str, updated_description: str = None) -> str:
    """Update a specific item in a workspace"""
    try:
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
def create_deployment_pipeline(prompt: str, pipeline_name: str = None, description: str = None) -> str:
    """
    Create a deployment pipeline with stages parsed from natural language
    
    Args:
        prompt: Natural language description of the pipeline (e.g., "Create a deployment pipeline with Dev and Test stages")
        pipeline_name: Optional custom name for the pipeline
        description: Optional description for the pipeline
    """
    try:
        stages = parse_stages_from_prompt(prompt)
        
        if not pipeline_name:
            stage_names = [stage['displayName'] for stage in stages]
            pipeline_name = f"Pipeline-{'-'.join(stage_names[:3])}"
        
        result = fabric_client.create_deployment_pipeline(
            display_name=pipeline_name,
            description=description,
            stages=stages
        )
        
        response_text = f"✅ Deployment pipeline '{pipeline_name}' created successfully!\n"
        response_text += f"ID: {result['id']}\n"
        response_text += f"Stages ({len(stages)}):\n"
        
        for i, stage in enumerate(stages):
            response_text += f"  {i+1}. {stage['displayName']}\n"
        
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
        response = fabric_client.delete_deployment_pipeline(pipeline_id)
        return f"Deployment pipeline deleted successfully"
    except Exception as e:
        return f"Error deleting deployment pipeline: {str(e)}"


@mcp.tool()
def delete_deployment_pipeline_role_assignment(pipeline_id: str, principal_id: str):
    try:
        response = fabric_client.delete_deployment_pipeline_role_assignment(pipeline_id, principal_id)
        return f"Role Assignment deleted successfully"
    except Exception as e:
        return f"Error deleting deployment pipeline role assignment: {str(e)}"


@mcp.tool()
def deploy_stage_content(pipeline_id: str, source_stage_id: str, target_stage_id: str, note: str = None) -> str:
    """Deploy content to a specific stage in a deployment pipeline"""
    try:
        response = fabric_client.deploy_stage_content(pipeline_id, source_stage_id, target_stage_id, note)
        return f"Stage content deployed successfully: {response}"
    except Exception as e:
        return f"Error deploying stage content: {str(e)}"


@mcp.tool()
def get_deployment_pipeline(pipeline_id: str) -> str:
    """Get details of a specific deployment pipeline"""
    try:
        response = fabric_client.get_deployment_pipeline(pipeline_id)
        return f"Deployment pipeline details: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline details: {str(e)}"


@mcp.tool()
def get_deployment_pipeline_stage(pipeline_id: str, stage_id: str) -> str:
    """Get details of a specific stage in a deployment pipeline"""
    try:
        response = fabric_client.get_deployment_pipeline_stage(pipeline_id, stage_id)
        return f"Deployment pipeline stage details: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline stage details: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_operations(pipeline_id: str) -> str:
    """List all operations for a deployment pipeline"""
    try:
        response = fabric_client.list_deployment_pipeline_operations(pipeline_id)
        return f"Deployment pipeline operations: {response}"
    except Exception as e:
        return f"Error listing deployment pipeline operations: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_role_assignments(pipeline_id: str) -> str:
    """List all role assignments for a deployment pipeline"""
    try:
        response = fabric_client.list_deployment_pipeline_role_assignments(pipeline_id)
        return f"Deployment pipeline role assignments: {response}"
    except Exception as e:
        return f"Error listing role assignments: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_stage_items(pipeline_id: str, stage_id: str) -> str:
    """List all items in a specific stage of a deployment pipeline"""
    try:
        response = fabric_client.list_deployment_pipeline_stage_items(pipeline_id, stage_id)
        return f"Deployment pipeline stage items: {response}"
    except Exception as e:
        return f"Error listing deployment pipeline stage items: {str(e)}"


@mcp.tool()
def list_deployment_pipeline_stages(pipeline_id: str, stage_id: str) -> str:
    """List stages of a specific deployment pipeline"""
    try:
        response = fabric_client.list_deployment_pipeline_stages(pipeline_id, stage_id)
        return f"Deployment pipeline stages: {response}"
    except Exception as e:
        return f"Error getting deployment pipeline stages: {str(e)}"


@mcp.tool()
def list_deployment_pipelines() -> str:
    """List all deployment pipelines"""
    try:
        response = fabric_client.list_deployment_pipelines()
        return f"Deployment pipelines: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing deployment pipelines: {str(e)}"


@mcp.tool()
def unassign_workspace_from_deployment_pipeline_stage(pipeline_id: str, stage_id: str) -> str:
    """Unassign a workspace from a specific stage in a deployment pipeline"""
    try:
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
        response = fabric_client.update_deployment_pipeline_stage(pipeline_id, stage_id, display_name, description)
        return f"Deployment pipeline stage updated successfully: {response}"
    except Exception as e:
        return f"Error updating deployment pipeline stage: {str(e)}"


@mcp.tool()
def assign_workspace_to_domain_by_capacity(domain_id: str, capacity_id: str) -> str:
    """Assign a workspace to a domain by capacity"""
    try:
        response = fabric_client.assign_workspace_to_domain_by_capacity(domain_id, capacity_id)
        return f"Workspace assigned to domain by capacity successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to domain: {str(e)}"


@mcp.tool()
def assign_workspace_to_domain_by_ids(domain_id: str, workspace_ids: List[str]) -> str:
    """Assign workspaces to a domain by IDs"""
    try:
        response = fabric_client.assign_workspace_to_domain_by_ids(domain_id, workspace_ids)
        return f"Workspaces assigned to domain successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to domain: {str(e)}"


@mcp.tool()
def create_domain(domain_name: str, description: str = None, parent_id: str = None) -> str:
    """Create a new domain in Microsoft Fabric"""
    try:
        response = fabric_client.create_domain(domain_name, description, parent_id)
        return f"Domain created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating domain: {str(e)}"


@mcp.tool()
def delete_domain(domain_id: str) -> str:
    """Delete a domain in Microsoft Fabric"""
    try:
        response = fabric_client.delete_domain(domain_id)
        return f"Domain deleted successfully"
    except Exception as e:
        return f"Error deleting domain: {str(e)}"


@mcp.tool()
def get_domain(domain_id: str) -> str:
    """Get details of a specific domain"""
    try:
        response = fabric_client.get_domain(domain_id)
        return f"Domain details: {response}"
    except Exception as e:
        return f"Error getting domain details: {str(e)}"


@mcp.tool()
def list_domain_workspaces(domain_id: str) -> str:
    """List all workspaces in a specific domain"""
    try:
        response = fabric_client.list_domain_workspaces(domain_id)
        return f"Domain workspaces: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing domain workspaces: {str(e)}"


@mcp.tool()
def list_domains() -> str:
    """List all domains in Microsoft Fabric"""
    try:
        response = fabric_client.list_domains()
        return f"Domains: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing domains: {str(e)}"


@mcp.tool()
def domain_bulk_assign_roles(domain_id: str, role: str, principals: List[Dict]) -> str:
    """Bulk assign roles to principals in a domain"""
    try:
        response = fabric_client.domain_bulk_assign_roles(domain_id, role, principals)
        return f"Roles bulk assigned successfully: {response}"
    except Exception as e:
        return f"Error bulk assigning roles in domain: {str(e)}"


@mcp.tool()
def domain_bulk_unassign_roles(domain_id: str, role: str, principals: List[Dict]) -> str:
    """Bulk unassign roles from principals in a domain"""
    try:
        response = fabric_client.domain_bulk_unassign_roles(domain_id, role, principals)
        return f"Roles bulk unassigned successfully: {response}"
    except Exception as e:
        return f"Error bulk unassigning roles in domain: {str(e)}"


@mcp.tool()
def domain_unassign_all_workspaces(domain_id: str) -> str:
    """Unassign all workspaces from a domain"""
    try:
        response = fabric_client.domain_unassign_all_workspaces(domain_id)
        return f"All workspaces unassigned from domain successfully: {response}"
    except Exception as e:
        return f"Error unassigning all workspaces from domain: {str(e)}"


@mcp.tool()
def domain_unassign_workspace_by_ids(domain_id: str, workspace_ids: List[str]) -> str:
    """Unassign specific workspaces from a domain"""
    try:
        response = fabric_client.domain_unassign_workspace_by_ids(domain_id, workspace_ids)
        return f"Workspaces unassigned from domain successfully: {response}"
    except Exception as e:
        return f"Error unassigning workspaces from domain: {str(e)}"


@mcp.tool()
def update_domain(domain_id: str, display_name: str = None, description: str = None) -> str:
    """Update an existing domain"""
    try:
        response = fabric_client.update_domain(domain_id, display_name, description)
        return f"Domain updated successfully: {response}"
    except Exception as e:
        return f"Error updating domain: {str(e)}"


@mcp.tool()
def apply_tags_to_item(workspace_id: str, item_id: str, tags_ids: List[str]) -> str:
    """Apply tags to a specific item"""
    try:
        result = fabric_client.apply_tags_to_item(workspace_id=workspace_id, item_id=item_id, tags=tags_ids)
        return f"✅ Tags applied to item '{item_id}' successfully!\nTags: {', '.join(tags_ids)}"
    except Exception as e:
        return f"❌ Error applying tags to item: {str(e)}"


@mcp.tool()
def list_tags_in_tenant() -> str:
    """List tags in the tenant"""
    try:
        response = fabric_client.list_tags_in_tenant()
        return f"Tenant tags: {response}"
    except Exception as e:
        return f"Error listing tags in tenant: {str(e)}"


@mcp.tool()
def unapply_tags_from_item(workspace_id: str, item_id: str, tags_ids: List[str]) -> str:
    """Remove tags from a specific item"""
    try:
        response = fabric_client.unapply_tags_from_item(workspace_id, item_id, tags_ids)
        return f"Tags removed successfully: {response}"
    except Exception as e:
        return f"Error unapplying tags from item: {str(e)}"


@mcp.tool()
def list_capacities() -> str:
    """List all capacities"""
    try:
        response = fabric_client.list_capacities()
        return f"Capacities: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing capacities: {str(e)}"


@mcp.tool()
def add_workspace_role_assignment(workspace_id: str, principal_id: str, principal_type: str, role: str) -> str:
    """Add a role assignment to a workspace"""
    try:
        response = fabric_client.add_workspace_role_assignment(workspace_id, principal_id, principal_type, role)
        return f"Workspace role assignment added successfully: {response}"
    except Exception as e:
        return f"Error adding workspace role assignment: {str(e)}"


@mcp.tool()
def assign_workspace_to_capacity(workspace_id: str, capacity_id: str) -> str:
    """Assign a workspace to a capacity"""
    try:
        response = fabric_client.assign_workspace_to_capacity(workspace_id, capacity_id)
        return f"Workspace assigned to capacity successfully: {response}"
    except Exception as e:
        return f"Error assigning workspace to capacity: {str(e)}"


@mcp.tool()
def create_workspace(display_name: str, description: str = None, capacity_id: str = None) -> str:
    """Create a new workspace"""
    try:
        response = fabric_client.create_workspace(display_name, description, capacity_id)
        return f"Workspace created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating workspace: {str(e)}"


@mcp.tool()
def delete_workspace(workspace_id: str) -> str:
    """Delete a workspace"""
    try:
        fabric_client.delete_workspace(workspace_id)
        return f"Workspace deleted successfully"
    except Exception as e:
        return f"Error deleting workspace: {str(e)}"


@mcp.tool()
def delete_workspace_role_assignment(workspace_id: str, role_assignment_id: str) -> str:
    """Delete a role assignment from a workspace"""
    try:
        fabric_client.delete_workspace_role_assignment(workspace_id, role_assignment_id)
        return f"Workspace role assignment deleted successfully"
    except Exception as e:
        return f"Error deleting workspace role assignment: {str(e)}"


@mcp.tool()
def get_workspace_by_id(workspace_id: str) -> str:
    """Get a workspace by its ID"""
    try:
        response = fabric_client.get_workspace_by_id(workspace_id)
        return f"Workspace details: {response}"
    except Exception as e:
        return f"Error getting workspace by ID: {str(e)}"


@mcp.tool()
def get_workspace_role_assignment(workspace_id: str, role_assignment_id: str) -> str:
    """Get a role assignment from a workspace"""
    try:
        response = fabric_client.get_workspace_role_assignment(workspace_id, role_assignment_id)
        return f"Workspace role assignment: {response}"
    except Exception as e:
        return f"Error getting workspace role assignment: {str(e)}"


@mcp.tool()
def list_workspace_role_assignments(workspace_id: str) -> str:
    """List all role assignments in a specific workspace"""
    try:
        response = fabric_client.list_workspace_role_assignments(workspace_id)
        return f"Workspace role assignments: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing workspace role assignments: {str(e)}"


@mcp.tool()
def list_workspaces() -> str:
    """List all workspaces"""
    try:
        response = fabric_client.list_workspaces()
        return f"Workspaces: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing workspaces: {str(e)}"


@mcp.tool()
def unassign_workspace_from_capacity(workspace_id: str) -> str:
    """Unassign a workspace from capacity"""
    try:
        fabric_client.unassign_workspace_from_capacity(workspace_id)
        return f"Workspace unassigned from capacity successfully"
    except Exception as e:
        return f"Error unassigning workspace from capacity: {str(e)}"


@mcp.tool()
def update_workspace(workspace_id: str, display_name: str = None, description: str = None) -> str:
    """Update a workspace"""
    try:
        response = fabric_client.update_workspace(workspace_id, display_name, description)
        return f"Workspace updated successfully: {response}"
    except Exception as e:
        return f"Error updating workspace: {str(e)}"


@mcp.tool()
def update_workspace_role_assignment(workspace_id: str, role_assignment_id: str, role: str = None) -> str:
    """Update a workspace role assignment"""
    try:
        response = fabric_client.update_workspace_role_assignment(workspace_id, role_assignment_id, role)
        return f"Workspace role assignment updated successfully: {response}"
    except Exception as e:
        return f"Error updating workspace role assignment: {str(e)}"


@mcp.tool()
def create_folder(workspace_id: str, folder_name: str, parent_folder_id: str = None) -> str:
    """Create a new folder in a specific workspace"""
    try:
        response = fabric_client.create_folder(workspace_id, folder_name, parent_folder_id)
        return f"Folder created successfully: {response.get('id', 'Unknown ID')}"
    except Exception as e:
        return f"Error creating folder in workspace: {str(e)}"


@mcp.tool()
def delete_folder(workspace_id: str, folder_id: str) -> str:
    """Delete a folder in a specific workspace"""
    try:
        fabric_client.delete_folder(workspace_id, folder_id)
        return f"Folder deleted successfully"
    except Exception as e:
        return f"Error deleting folder in workspace: {str(e)}"


@mcp.tool()
def get_folder(workspace_id: str, folder_id: str) -> str:
    """Get details of a specific folder in a workspace"""
    try:
        response = fabric_client.get_folder(workspace_id, folder_id)
        return f"Folder details: {response}"
    except Exception as e:
        return f"Error getting folder in workspace: {str(e)}"


@mcp.tool()
def list_folders(workspace_id: str, root_folder_id: str = None, recursive: bool = False) -> str:
    """List all folders in a specific workspace"""
    try:
        response = fabric_client.list_folders(workspace_id, root_folder_id, recursive)
        return f"Folders: {len(response)} found - {response}"
    except Exception as e:
        return f"Error listing folders in workspace: {str(e)}"


@mcp.tool()
def move_folder(workspace_id: str, folder_id: str, target_folder_id: str) -> str:
    """Move a folder to a new parent folder in a specific workspace"""
    try:
        fabric_client.move_folder(workspace_id, folder_id, target_folder_id)
        return f"Folder moved successfully"
    except Exception as e:
        return f"Error moving folder in workspace: {str(e)}"


@mcp.tool()
def update_folder(workspace_id: str, folder_id: str, new_folder_name: str) -> str:
    """Update the name of a folder in a specific workspace"""
    try:
        fabric_client.update_folder(workspace_id, folder_id, new_folder_name)
        return f"Folder updated successfully"
    except Exception as e:
        return f"Error updating folder in workspace: {str(e)}"


@mcp.tool()
def analyze_semantic_model_dependencies(workspace_id: str, dataset_id: str) -> str:
    """
    Analyze Power BI semantic model (dataset) dependencies and create lineage data
    
    This tool analyzes the dependency chain of a Power BI semantic model, finding all datasets 
    and reports that depend on the target dataset, creating a comprehensive lineage view.
    
    Args:
        workspace_id: The Power BI workspace ID (GUID format)
        dataset_id: The ID of the target semantic model/dataset to analyze (GUID format)
        
    Returns:
        Detailed dependency analysis with lineage data
    """
    try:
        result = fabric_client.analyze_semantic_model_dependencies(workspace_id, dataset_id)
        
        if not result.get("success"):
            return f"Analysis failed: {result}"
        
        # Build clean response without excessive formatting
        lines = [
            "Semantic Model Dependency Analysis Complete",
            "=" * 50,
            f"Target Dataset: {result['target_dataset_name']}",
            f"Dataset ID: {result['target_dataset_id']}",
            f"Workspace ID: {result['workspace_id']}",
            "",
            f"Analysis Summary:",
            f"- Total Datasets Analyzed: {result['total_datasets_analyzed']}",
            f"- Dependent Datasets Found: {result['dependent_datasets']}",
            f"- Reports Found: {result['total_reports']}",
            "",
            "Lineage Details:"
        ]
        
        # Group by role for better organization
        main_items = [item for item in result['lineage_data'] if item['role'] == 'Main']
        dependent_items = [item for item in result['lineage_data'] if item['role'] == 'Dependent']
        
        # Add main dataset info
        if main_items:
            lines.append("\nMain Dataset:")
            for item in main_items:
                if item['report_name']:
                    lines.append(f"  Dataset: {item['dataset_name']}")
                    lines.append(f"    Report: {item['report_name']}")
                else:
                    lines.append(f"  Dataset: {item['dataset_name']} (No Reports)")
        
        # Add dependent datasets
        if dependent_items:
            lines.append("\nDependent Datasets:")
            current_dataset = None
            for item in dependent_items:
                if item['dataset_name'] != current_dataset:
                    current_dataset = item['dataset_name']
                    lines.append(f"  Dataset: {item['dataset_name']}")
                
                if item['report_name']:
                    lines.append(f"    Report: {item['report_name']}")
            
            # Add summary for datasets without reports
            datasets_without_reports = [item for item in dependent_items if not item['report_name']]
            if datasets_without_reports:
                unique_datasets = set(item['dataset_name'] for item in datasets_without_reports)
                if unique_datasets:
                    lines.append(f"\nDatasets without reports: {len(unique_datasets)}")
        
        if not dependent_items:
            lines.append("\nNo dependent datasets found.")
        
        lines.extend([
            "",
            "Analysis completed successfully.",
            "This data can be used for impact analysis and deployment planning."
        ])
        
        return "\n".join(lines)
            
    except Exception as e:
        return f"Error analyzing semantic model dependencies: {str(e)}"



@mcp.tool()
def check_config() -> str:
    """Check if the authentication configuration is properly set up"""
    config_status = []
    
    if fabric_client.tenant_id:
        config_status.append(f"✅ FABRIC_TENANT_ID is set")
    else:
        config_status.append("❌ FABRIC_TENANT_ID is missing")
    
    if fabric_client.client_id:
        config_status.append(f"✅ FABRIC_CLIENT_ID is set")
    else:
        config_status.append("❌ FABRIC_CLIENT_ID is missing")
    
    if fabric_client.client_secret:
        config_status.append(f"✅ FABRIC_CLIENT_SECRET is set")
    else:
        config_status.append("❌ FABRIC_CLIENT_SECRET is missing")
    
    all_configured = all([fabric_client.tenant_id, fabric_client.client_id, fabric_client.client_secret])
    
    status_text = "\n".join(config_status)
    if all_configured:
        status_text += "\n\n✅ Configuration is complete and ready to use!"
    else:
        status_text += "\n\n❌ Configuration is incomplete. Please set the missing environment variables."
    
    return status_text

mcp_app = mcp.http_app()

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    Ensure all services, including MCP, are initialized before the app starts serving.
    """
    global initialization_complete
    async with mcp_app.lifespan(app) as mcp_lifespan_manager:
        print("INFO:     MCP App startup complete.")
        yield
        print("INFO:     MCP App shutdown starting.")
    print("INFO:     MCP App shutdown complete.")

app = FastAPI(lifespan=lifespan)

app.mount("/mcp-server", mcp_app)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)