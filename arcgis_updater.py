import os
import json
import requests
from datetime import datetime
import traceback
import logging

class ArcGISUpdater:
    def __init__(self):
        # Load credentials from environment variables
        self.base_url = os.getenv('ARCGIS_URL')
        self.username = os.getenv('ARCGIS_USERNAME')
        self.password = os.getenv('ARCGIS_PASSWORD')
        
        if not all([self.base_url, self.username, self.password]):
            raise ValueError("Missing required ArcGIS Enterprise credentials in environment variables")
        
        # Get token
        self.token = self._get_token()
        logging.info(f"Successfully connected to ArcGIS Enterprise as {self.username}")
        
        # Store feature service URLs using environment variables for layer IDs
        self.feature_services = {
            'poles': f"{self.base_url}/{os.getenv('ARCGIS_POLES_LAYER_ID', '1')}",
            'connections': f"{self.base_url}/{os.getenv('ARCGIS_CONNECTIONS_LAYER_ID', '3')}",
            'anchors': f"{self.base_url}/{os.getenv('ARCGIS_ANCHORS_LAYER_ID', '2')}"
        }
        
        logging.info("Feature service URLs configured:")
        for layer, url in self.feature_services.items():
            logging.info(f"{layer}: {url}")
    
    def _get_token(self):
        """Get an authentication token from ArcGIS Enterprise Portal"""
        portal_token_url = "https://gis.clearnetworx.com/portal/sharing/rest/generateToken"
        
        params = {
            'username': self.username,  # Using brandan.lewis from env
            'password': self.password,
            'client': 'referer',
            'referer': 'https://www.arcgis.com',
            'f': 'json',
            'expiration': 60
        }
        
        print(f"Debug - Using credentials - Username: {self.username}")
        print(f"Debug - Token URL: {portal_token_url}")
        
        try:
            response = requests.post(portal_token_url, data=params, verify=False)
            response.raise_for_status()
            token_info = response.json()
            
            if "token" not in token_info:
                raise Exception(f"Failed to get portal token: {token_info}")
            
            return token_info["token"]
        except Exception as e:
            raise ValueError(f"Failed to get token: {str(e)}")
    
    def delete_features_by_job(self, layer_name, job_names):
        """Delete all features for specified jobs before adding new ones"""
        if layer_name not in self.feature_services:
            raise ValueError(f"Feature service for {layer_name} not found")
            
        service_url = self.feature_services[layer_name]
        
        try:
            # Build where clause to match job names
            where_clause = " OR ".join([f"job_name = '{job}'" for job in job_names])
            
            # Query for features to delete
            query_url = f"{service_url}/query"
            query_params = {
                'f': 'json',
                'token': self.token,
                'where': where_clause,
                'returnIdsOnly': 'true'
            }
            
            response = requests.post(query_url, data=query_params, verify=False)
            result = response.json()
            
            if 'objectIds' in result and result['objectIds']:
                object_ids = result['objectIds']
                print(f"Found {len(object_ids)} existing features to delete for {layer_name}")
                
                # Delete features in chunks
                chunk_size = 100
                for i in range(0, len(object_ids), chunk_size):
                    chunk = object_ids[i:i + chunk_size]
                    delete_params = {
                        'f': 'json',
                        'token': self.token,
                        'objectIds': ','.join(map(str, chunk))
                    }
                    
                    delete_url = f"{service_url}/deleteFeatures"
                    delete_response = requests.post(delete_url, data=delete_params, verify=False)
                    delete_result = delete_response.json()
                    
                    if 'error' in delete_result:
                        print(f"Error deleting features: {delete_result['error']}")
                    else:
                        print(f"Successfully deleted {len(chunk)} features")
            else:
                print(f"No existing features found for {layer_name} with jobs: {job_names}")
                
        except Exception as e:
            print(f"Error deleting features: {str(e)}")

    def clear_all_features(self, layer_name):
        """Delete all features from a layer"""
        if layer_name not in self.feature_services:
            raise ValueError(f"Feature service for {layer_name} not found")
            
        service_url = self.feature_services[layer_name]
        
        try:
            # Query for all features
            query_url = f"{service_url}/query"
            query_params = {
                'f': 'json',
                'token': self.token,
                'where': '1=1',  # Get all features
                'returnIdsOnly': 'true'
            }
            
            response = requests.post(query_url, data=query_params, verify=False)
            result = response.json()
            
            if 'objectIds' in result and result['objectIds']:
                object_ids = result['objectIds']
                print(f"Found {len(object_ids)} existing features to delete from {layer_name}")
                
                # Delete features in chunks
                chunk_size = 100
                for i in range(0, len(object_ids), chunk_size):
                    chunk = object_ids[i:i + chunk_size]
                    delete_params = {
                        'f': 'json',
                        'token': self.token,
                        'objectIds': ','.join(map(str, chunk))
                    }
                    
                    delete_url = f"{service_url}/deleteFeatures"
                    delete_response = requests.post(delete_url, data=delete_params, verify=False)
                    delete_result = delete_response.json()
                    
                    if 'error' in delete_result:
                        print(f"Error deleting features: {delete_result['error']}")
                    else:
                        print(f"Successfully deleted {len(chunk)} features")
            else:
                print(f"No existing features found in {layer_name}")
                
        except Exception as e:
            print(f"Error clearing features: {str(e)}")

    def update_features(self, layer_name, features, test_mode=True):
        """Add new features to the layer (after deleting existing ones)"""
        if layer_name not in self.feature_services:
            raise ValueError(f"Feature service for {layer_name} not found")
            
        service_url = self.feature_services[layer_name]
        total_features = len(features)
        logging.info(f"\nAttempting to add {total_features} {layer_name}")
        
        if layer_name == 'connections':
            # Log sample connection for debugging
            if features:
                sample = features[0]
                logging.info(f"\nSample connection feature:")
                logging.info(f"Geometry: {json.dumps(sample['geometry'], indent=2)}")
                logging.info(f"Attributes: {json.dumps(sample['attributes'], indent=2)}")
                logging.info(f"Wire spec present in {sum(1 for f in features if f['attributes'].get('wire_spec'))} features")
                logging.info(f"Connection types: {set(f['attributes'].get('conn_type', '') for f in features)}")
        
        try:
            # Test the service URL before proceeding
            test_url = f"{service_url}?f=json&token={self.token}"
            response = requests.get(test_url, verify=False)
            if response.status_code == 404:
                logging.error(f"Feature service not found: {service_url}")
                logging.error(f"Response: {response.text}")
                return False
            elif response.status_code != 200:
                logging.error(f"Error accessing feature service: {response.status_code}")
                logging.error(f"Response: {response.text}")
                return False
                
            service_info = response.json()
            logging.info(f"Successfully connected to feature service: {layer_name}")
            logging.info(f"Service info: {json.dumps(service_info, indent=2)}")
            
            # In test mode, clear all existing features
            if test_mode:
                logging.info(f"Test mode: Clearing all existing features from {layer_name}")
                self.clear_all_features(layer_name)
            
            # Map fields according to feature service schema
            mapped_features = []
            for feature in features:
                mapped_feature = {
                    'geometry': feature['geometry']
                }
                
                # Map attributes based on layer type
                attrs = feature.get('attributes', {})
                if layer_name == 'poles':
                    mapped_feature['attributes'] = {
                        'node_id': attrs.get('node_id', ''),
                        'job_name': attrs.get('job_name', ''),
                        'job_stat': attrs.get('job_status', ''),
                        'mr_status': attrs.get('mr_status', ''),
                        'utility': attrs.get('utility', ''),
                        'completed': attrs.get('completed', ''),
                        'pole_tag': attrs.get('pole_tag', ''),
                        'poa_ht': attrs.get('poa_ht', ''),
                        'conv': attrs.get('conv', ''),
                        'proj': attrs.get('proj', ''),
                        'scid': attrs.get('scid', ''),
                        'editor': attrs.get('last_editor', ''),
                        'edit_time': attrs.get('last_edit', '')
                    }
                elif layer_name == 'connections':
                    mapped_feature['attributes'] = {
                        'conn_id': attrs.get('conn_id', ''),
                        'conn_type': attrs.get('conn_type', ''),
                        'mid_ht': attrs.get('att_height', ''),
                        'node1_id': attrs.get('node_id_1', ''),
                        'node2_id': attrs.get('node_id_2', ''),
                        'wire_spec': attrs.get('wire_spec', '')
                    }
                elif layer_name == 'anchors':
                    mapped_feature['attributes'] = {
                        'anch_spec': attrs.get('anch_spec', ''),
                        'job_id': attrs.get('job_id', ''),
                        'anchor_typ': attrs.get('anchor_type', '')
                    }
                
                mapped_features.append(mapped_feature)
            
            # Get unique job names from features
            job_names = set()
            for feature in mapped_features:
                job_name = feature.get('attributes', {}).get('job_name', '')
                if job_name:
                    job_names.add(job_name)
            
            # Delete existing features for these jobs
            if job_names:
                logging.info(f"Deleting existing features for jobs: {job_names}")
                self.delete_features_by_job(layer_name, job_names)
            
            # Add features in chunks
            chunk_size = 100
            successful_adds = 0
            failed_adds = 0
            
            for i in range(0, len(mapped_features), chunk_size):
                chunk = mapped_features[i:i + chunk_size]
                chunk_start = i + 1
                chunk_end = min(i + chunk_size, total_features)
                
                try:
                    # Print sample of features being added
                    sample_feature = chunk[0]
                    logging.info(f"\nAdding {layer_name} {chunk_start}-{chunk_end} of {total_features}")
                    logging.info(f"Sample feature: {json.dumps(sample_feature, indent=2)}")
                    
                    # Prepare the add request
                    add_url = f"{service_url}/addFeatures"
                    params = {
                        'f': 'json',
                        'token': self.token,
                        'features': json.dumps(chunk)
                    }
                    
                    response = requests.post(add_url, data=params, verify=False)
                    response.raise_for_status()
                    result = response.json()
                    
                    if 'error' in result:
                        logging.error(f"Error adding {layer_name} features {chunk_start}-{chunk_end}:")
                        logging.error(json.dumps(result, indent=2))
                        failed_adds += len(chunk)
                    else:
                        logging.info(f"Successfully added {len(chunk)} {layer_name} features ({chunk_start}-{chunk_end})")
                        successful_adds += len(chunk)
                        
                except Exception as e:
                    logging.error(f"Error adding {layer_name} chunk {chunk_start}-{chunk_end}: {str(e)}")
                    logging.error(f"Stack trace: {traceback.format_exc()}")
                    failed_adds += len(chunk)
                    # Try to refresh token if it might have expired
                    if 'token' in str(e).lower():
                        try:
                            self.token = self._get_token()
                            logging.info("Token refreshed, retrying add...")
                            params['token'] = self.token
                            response = requests.post(add_url, data=params, verify=False)
                            response.raise_for_status()
                            result = response.json()
                            if 'error' not in result:
                                logging.info(f"Successfully added {len(chunk)} {layer_name} features after token refresh")
                                successful_adds += len(chunk)
                                failed_adds -= len(chunk)
                        except Exception as retry_error:
                            logging.error(f"Error retrying add after token refresh: {str(retry_error)}")
            
            # Print summary for this layer
            logging.info(f"\nSummary for {layer_name}:")
            logging.info(f"Total features attempted: {total_features}")
            logging.info(f"Successfully added: {successful_adds}")
            logging.info(f"Failed adds: {failed_adds}")
            logging.info(f"Success rate: {(successful_adds/total_features*100 if total_features > 0 else 0):.1f}%")
            logging.info("-" * 50)
            
            return successful_adds > 0
            
        except Exception as e:
            logging.error(f"Error updating {layer_name}: {str(e)}")
            logging.error(f"Stack trace: {traceback.format_exc()}")
            return False
    
    def process_shapefile(self, shapefile_path, layer_name):
        """Process a shapefile and update the corresponding feature service"""
        try:
            print(f"\nProcessing {layer_name} shapefile: {shapefile_path}")
            
            # Read shapefile
            gdf = gpd.read_file(shapefile_path)
            print(f"Read {len(gdf)} features from shapefile")
            print(f"Columns found: {', '.join(gdf.columns)}")
            
            # Print sample of the data
            print("\nSample of first feature:")
            sample_row = gdf.iloc[0]
            for col in gdf.columns:
                if col != 'geometry':
                    print(f"{col}: {sample_row[col]}")
            print(f"Geometry type: {sample_row.geometry.geom_type}")
            
            # Convert features to ArcGIS JSON format
            features = []
            invalid_geoms = 0
            for idx, row in gdf.iterrows():
                try:
                    # Create the feature with proper field mappings based on layer type
                    if layer_name == 'poles':
                        job_name = str(row.get('job_name', ''))  # Get job name from poles
                        attributes = {
                            'node_id': str(row.get('node_id', '')),
                            'job_name': job_name,
                            'job_stat': str(row.get('job_stat', '')),
                            'mr_status': str(row.get('mr_status', '')),
                            'utility': str(row.get('utility', '')),
                            'completed': str(row.get('completed', '')),
                            'pole_tag': str(row.get('pole_tag', '')),
                            'scid': str(row.get('scid', '')),
                            'poa_ht': str(row.get('poa_ht', '')),
                            'conv': str(row.get('conv', '')),
                            'proj': str(row.get('proj', '')),
                            'editor': str(row.get('editor', '')),
                            'edit_time': str(row.get('edit_time', ''))
                        }
                    elif layer_name == 'connections':
                        # Get job name from the first node's job name
                        node1_id = str(row.get('node_id_1', ''))
                        job_name = str(row.get('job_name', ''))  # Try to get from shapefile first
                        attributes = {
                            'connection_id': str(row.get('connection_id', '')),
                            'job_name': job_name,  # Add job_name for connections
                            'connection_type': str(row.get('connection_type', '')),
                            'wire_spec': str(row.get('wire_spec', '')),
                            'mid_ht': str(row.get('mid_ht', '')),
                            'node_id_1': node1_id,
                            'node_id_2': str(row.get('node_id_2', ''))
                        }
                    elif layer_name == 'anchors':
                        job_name = str(row.get('job_name', ''))  # Try to get from shapefile first
                        if not job_name:
                            job_id = str(row.get('job_id', ''))
                            job_name = f"Job {job_id}"  # Create job name from job_id if needed
                        attributes = {
                            'anch_spec': str(row.get('anchor_spec', '')),
                            'anchor_typ': str(row.get('anchor_type', '')),
                            'job_id': str(row.get('job_id', '')),
                            'job_name': job_name  # Add job_name for anchors
                        }
                    else:
                        raise ValueError(f"Unknown layer type: {layer_name}")

                    # Print debug info about job name
                    print(f"Processing {layer_name} feature {idx} for job: {job_name}")

                    feature = {
                        'geometry': json.loads(row.geometry.json),
                        'attributes': attributes
                    }
                    features.append(feature)
                except Exception as e:
                    print(f"Error processing feature {idx}: {str(e)}")
                    invalid_geoms += 1
            
            print(f"\nFeature conversion summary:")
            print(f"Total features processed: {len(gdf)}")
            print(f"Valid features: {len(features)}")
            print(f"Invalid geometries: {invalid_geoms}")
            
            if features:
                # Group features by job name for better logging
                job_counts = {}
                for feature in features:
                    job = feature['attributes'].get('job_name', 'Unknown')
                    job_counts[job] = job_counts.get(job, 0) + 1
                print("\nFeatures by job:")
                for job, count in job_counts.items():
                    print(f"{job}: {count} features")
                
                # Update features
                self.update_features(layer_name, features)
            else:
                print(f"No valid features to update for {layer_name}")
            
        except Exception as e:
            print(f"Error processing shapefile {shapefile_path}: {str(e)}")
            print("Stack trace:", traceback.format_exc())
    
    def process_master_zip(self, master_zip_path):
        """Process all layers in the master zip file"""
        import tempfile
        import zipfile
        
        with zipfile.ZipFile(master_zip_path, 'r') as zip_ref:
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_ref.extractall(tmpdir)
                
                # Process each layer
                layer_files = {
                    'poles': 'poles.shp',
                    'connections': 'connections.shp',
                    'anchors': 'anchors.shp'
                }
                
                for layer_name, filename in layer_files.items():
                    shapefile_path = os.path.join(tmpdir, filename)
                    if os.path.exists(shapefile_path):
                        print(f"\nProcessing {layer_name}...")
                        self.process_shapefile(shapefile_path, layer_name)
                    else:
                        print(f"Warning: {filename} not found in master zip") 