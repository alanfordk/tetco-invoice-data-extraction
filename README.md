# TETCO Invoice Data Collection & Automation

This project was developed to help digitize and analyze 25+ years of invoice data for **TETCO**, a small environmental emissions testing company operating in the Intermountain West for over 30 years.

## Project Objective

The goal was to extract key business information from decades of unstructured invoice files (PDF, DOCX, and WPD formats) and convert it into a structured format for analysis. The final dataset supports metrics such as:

- Revenue by client and year
- Frequency of repeat customers
- Job location tracking
- Estimation of costs by service type
- Visualization of company trends and growth

## What This Project Does

- Reads and parses invoice files from multiple formats:
  - `.pdf` using **pdfplumber**
  - `.docx` using **python-docx**
  - `.wpd` (WordPerfect) using **Apache Tika**
- Extracts key fields such as:
  - Client name
  - Project description
  - Start and end test dates
  - Invoice amount and date
  - Location
- Appends extracted data to a structured Excel spreadsheet

## Technologies Used

- Python
- pdfplumber
- python-docx
- Apache Tika
- pandas
- openpyxl

## Use of AI

Many of the scripts were initially developed with the assistance of AI tools like ChatGPT. I used these tools to help brainstorm approaches, debug errors, and write boilerplate code — but I tested and customized everything myself to make it work for our specific document formats and goals.

This project helped me learn:
- How to integrate AI coding support into real-world problem-solving
- How to work with messy, inconsistent data
- How to automate repetitive tasks with real business impact

## Business Impact

By automating the invoice data collection process, what would have taken **months of manual data entry** was completed in **under a week**. This provided the company with:
- A searchable record of historical jobs
- Better understanding of customer behavior
- Tools to inform future pricing, strategy, and bidding

##  About Me

I'm currently transitioning into a career in data and automation. This project gave me hands-on experience working with messy, real-world data and solving problems that bring value to a business. Feel free to connect with me on [LinkedIn](https://www.linkedin.com/in/alan-kitchen).

---

*This repository is for learning and demonstration purposes only. Client names and financial information have been anonymized or excluded.*
