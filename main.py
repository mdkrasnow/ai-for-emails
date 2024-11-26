from typing import List, Dict, Optional
import pandas as pd
from openai import OpenAI
from pathlib import Path
import os
from dotenv import load_dotenv
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Configure OpenAI
api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    raise ValueError("OpenAI API key not found in environment variables")

client = OpenAI(api_key=api_key)


class EmailCustomizer:
    def __init__(self, file_path: str, generic_email: str):
        self.file_path = Path(file_path)
        self.generic_email = generic_email
        self.df: Optional[pd.DataFrame] = None
        
        if not self.file_path.exists():
            raise FileNotFoundError(f"csv file not found at {file_path}")
            
    def read_csv(self) -> None:
        """Read the csv file into a pandas DataFrame."""
        try:
            self.df = pd.read_csv(self.file_path)
            logger.info(f"Successfully read csv file with {len(self.df)} rows")
        except Exception as e:
            logger.error(f"Error reading csv file: {str(e)}")
            raise

    def generate_custom_email(self, company_name: str, company_description: str) -> str:
        """Generate a customized email using OpenAI API."""
        try:
            prompt = f"""
            Original email template:
            {self.generic_email}
            
            Company name:
            {company_name}
            
            Company description:
            {company_description}
            
            Please customize the email template for this specific company,
            incorporating relevant details from the company name and description.
            Keep the email professional and concise.
            You will only output the customized email content.
            """
            
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a student applying for an internship. You maintain a professional tone."},
                    {"role": "user", "content": prompt},
                ],
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            logger.error(f"Error generating custom email: {str(e)}")
            return f"Error generating email: {str(e)}"

    def process_spreadsheet(self) -> None:
        """Process the entire spreadsheet and generate custom emails."""
        if self.df is None:
            self.read_csv()
            
        # Get the first column name (company name) and fourth column name (description)
        if len(self.df.columns) < 4:
            raise ValueError("Spreadsheet must have at least 4 columns")
            
        company_name_column = self.df.columns[0]  # Get first column name
        description_column = self.df.columns[3]  # Get fourth column name
        output_column = self.df.columns[4] if len(self.df.columns) > 4 else 'Generated_Email'
            
        # Process each row
        for index, row in self.df.iterrows():
            company_name = row[company_name_column]
            company_description = row[description_column]
            if pd.isna(company_description) or pd.isna(company_name):
                continue
                
            logger.info(f"Processing row {index + 1}")
            custom_email = self.generate_custom_email(
                str(company_name),
                str(company_description)
            )
            
            # Update the output column with the custom email
            self.df.at[index, output_column] = custom_email
            
        # Save the updated spreadsheet
        self.save_spreadsheet()
        
    def save_spreadsheet(self) -> None:
        """Save the updated spreadsheet."""
        try:
            output_path = self.file_path.with_name(f"{self.file_path.stem}_updated{self.file_path.suffix}")
            self.df.to_csv(output_path, index=False)
            logger.info(f"Successfully saved updated spreadsheet to {output_path}")
        except Exception as e:
            logger.error(f"Error saving spreadsheet: {str(e)}")
            raise

def main():
    try:
        generic_email = """
        Dear Hiring Team,

        My name is Juliet Sostena, and I am currently studying Bioengineering at Stanford University. I am very excited about the work your company is doing, as it closely aligns with my research background and academic focus. Over the last four years, I've gained extensive wet-lab and computational research experience in neurodegenerative diseases, brain aging, protein interactions, and more. I have also gained experience in the business administration side of biotech through recent internships. I would love to learn more about potential summer internship opportunities within your organization.

        Thank you for your time, and I look forward to hearing from you soon.

        Best regards,

        Juliet Sostena
        """
        # Get input from user
        file_path = "spreadsheet.csv"
        generic_email = generic_email

        customizer = EmailCustomizer(file_path, generic_email)
        
        customizer.process_spreadsheet()
        
        logger.info("Spreadsheet customization complete")
        
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
        raise

if __name__ == "__main__":
    main()