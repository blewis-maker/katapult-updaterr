import schedule
import time
from datetime import datetime
import logging
from incremental_update import daily_update, weekly_update

# Configure logging
logging.basicConfig(
    filename='logs/scheduler.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def run_daily_update():
    """Run the daily update at 11:59 PM"""
    try:
        logging.info("Starting daily update...")
        metrics = daily_update()
        logging.info(f"Daily update completed successfully. Processed {metrics['total_jobs']} jobs")
    except Exception as e:
        logging.error(f"Error in daily update: {str(e)}")

def run_weekly_update():
    """Run the weekly update on Saturday nights"""
    try:
        logging.info("Starting weekly update...")
        metrics = weekly_update()
        logging.info(f"Weekly update completed successfully. Processed {metrics['total_jobs']} jobs")
    except Exception as e:
        logging.error(f"Error in weekly update: {str(e)}")

def main():
    # Schedule daily update at 11:59 PM
    schedule.every().day.at("23:59").do(run_daily_update)
    
    # Schedule weekly update on Saturday at 11:59 PM
    schedule.every().saturday.at("23:59").do(run_weekly_update)
    
    logging.info("Scheduler started. Waiting for scheduled runs...")
    
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute

if __name__ == "__main__":
    main() 