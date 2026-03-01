#!/usr/bin/env python
# src/financial_researcher/main.py
import os
from financial_researcher.crew import ResearchCrew
from financial_researcher.ppt_generator import convert_report_to_ppt

# Create output directory if it doesn't exist
os.makedirs('output', exist_ok=True)

def run():
    """
    Run the research crew and generate PowerPoint presentation.
    """
    inputs = {
        'company': 'Apple'
    }

    # Create and run the crew
    result = ResearchCrew().crew().kickoff(inputs=inputs)

    # Print the result
    print("\n\n=== FINAL REPORT ===\n\n")
    print(result.raw)

    print("\n\nReport has been saved to output/report.md")
    
    # Generate PowerPoint from the report with company name
    try:
        company_name = inputs['company']
        ppt_path = convert_report_to_ppt('output/report.md', 'output/report.pptx', company_name=company_name)
        print(f"\n✓ PowerPoint presentation has been generated: {ppt_path}")
    except Exception as e:
        print(f"\n✗ Error generating PowerPoint: {e}")

if __name__ == "__main__":
    run()