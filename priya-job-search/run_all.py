from linkedin.linkedin_scraper import run_linkedin
from naukri.naukri_scraper import run_naukri
from indeed.indeed_scraper import run_indeed
from common.pipeline import combine_and_write
from common.config import load_config

def main():
    config = load_config()

    print("\n=== LOCAL JOB SEARCH ===\n")

    linkedin_jobs = []
    naukri_jobs = []
    indeed_jobs = []

    print("[1/2] LinkedIn")
    try:
        linkedin_jobs = run_linkedin(config)
        print(f"LinkedIn: {len(linkedin_jobs)} jobs captured")
    except Exception as e:
        print(f"LinkedIn failed: {e}")
        print("Continuing with Naukri...")

    print("\n[2/2] Naukri")
    try:
        naukri_jobs = run_naukri(config)
        print(f"Naukri: {len(naukri_jobs)} jobs captured")
    except Exception as e:
        print(f"Naukri failed: {e}")
        print("Continuing with available results...")

    print("\n[3/3] Indeed")
    try:
        indeed_jobs = run_indeed(config)
        print(f"Indeed: {len(indeed_jobs)} jobs captured")
    except Exception as e:
        print(f"Indeed failed: {e}")
        print("Continuing with available results...")

    combine_and_write(linkedin_jobs + naukri_jobs + indeed_jobs, config)

if __name__ == "__main__":
    main()
