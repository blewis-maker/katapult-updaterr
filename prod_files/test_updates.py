import os
import logging
import sys
from datetime import datetime
from incremental_update import process_daily_update
from main import (
    TEST_JOB_IDS,
    getJobData,
    extractNodes,
    extractConnections,
    extractAnchors,
    getUserList,
    update_arcgis_features,
    saveToShapefiles,
    CONFIG
)

# Ensure required directories exist in Docker environment
os.makedirs('/app/logs', exist_ok=True)
os.makedirs(CONFIG['WORKSPACE_PATH'], exist_ok=True)  # Ensure workspace directory exists

# Configure logging for Docker environment
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('/app/logs/test_updates.log')
    ]
)
logger = logging.getLogger(__name__)

def test_arcgis_update():
    """Test ArcGIS update functionality using test jobs"""
    logger.info("Starting ArcGIS update test with test jobs...")
    
    try:
        # Ensure we're in the correct directory
        os.chdir('/app')
        
        # Get user list for mapping IDs to names
        user_map = getUserList()
        logger.info(f"Retrieved {len(user_map)} users")
        
        all_nodes = []
        all_connections = []
        all_anchors = []
        jobs_summary = []
        
        # Process each test job
        total_jobs = len(TEST_JOB_IDS)
        logger.info(f"Found {total_jobs} test jobs to process")
        
        for index, job_id in enumerate(TEST_JOB_IDS, 1):
            logger.info(f"\nProcessing job {index}/{total_jobs}: {job_id}")
            
            # Get job data
            job_data = getJobData(job_id)
            if not job_data:
                logger.warning(f"No data found for job {job_id}")
                continue
            
            # Extract job name and metadata
            metadata = job_data.get('metadata', {})
            job_name = metadata.get('name', f"Job {job_id}")
            
            # Extract nodes with user map
            logger.info("Extracting nodes...")
            nodes_data = extractNodes(job_data, job_name, job_id, user_map)
            logger.info(f"Found {len(nodes_data)} nodes")
            
            if nodes_data:
                all_nodes.extend(nodes_data)
                
                # Extract connections with full job data
                logger.info("Extracting connections...")
                connections = job_data.get('connections', {})
                connections_data = extractConnections(connections, job_data.get('nodes', {}), job_data)
                if connections_data:
                    all_connections.extend(connections_data)
                    logger.info(f"Found {len(connections_data)} connections")
                
                # Extract anchors
                logger.info("Extracting anchors...")
                anchors = extractAnchors(job_data, job_name, job_id)
                if anchors:
                    all_anchors.extend(anchors)
                    logger.info(f"Found {len(anchors)} anchors")
            
            logger.info(f"Completed processing job {job_id}")
        
        if all_nodes or all_connections or all_anchors:
            # Update ArcGIS feature services
            logger.info("Updating ArcGIS feature services...")
            arcgis_success = update_arcgis_features(all_nodes, all_connections, all_anchors)
            if arcgis_success:
                logger.info("ArcGIS feature services updated successfully")
            else:
                logger.error("Failed to update ArcGIS feature services")
            
            # Save to shapefiles and update SharePoint
            workspace_path = CONFIG['WORKSPACE_PATH']
            logger.info("Saving data to shapefiles...")
            saveToShapefiles(all_nodes, all_connections, all_anchors, workspace_path)
            
            logger.info(f"Successfully processed {len(TEST_JOB_IDS)} test jobs")
        else:
            logger.info("No data extracted for any job. Nothing to save.")
            
        return True
        
    except Exception as e:
        logger.error(f"Error in ArcGIS update test: {str(e)}")
        return False

if __name__ == '__main__':
    # Run the ArcGIS update test
    success = test_arcgis_update()
    exit(0 if success else 1) 