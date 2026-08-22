from playwright.sync_api import sync_playwright
import time

def main():
    
    with sync_playwright() as playwright:
        user_data_dir = "chrome-profile"
        
        try:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir,
                headless=False,
                args=['--remote-debugging-port=9222']
            )
            
            page = context.pages[0] if context.pages else context.new_page()
            
            print("\nOpening NHA Provider portal...")
            page.goto("https://provider.nha.gov.in/")
            
          
            while True:
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n\nShutting down browser...")
            context.close()
            
        except Exception as e:
            print(f"\nError: {e}")
            print("  rm chrome-profile/SingletonLock")

if __name__ == "__main__":
    main()