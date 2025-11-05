# Pckgd
# Title: PDFCreator
# Description: A tool to test creating pdf's using iTextSharp
# Updates from: GitHub/bswaby/Touchpoint/Python%20Scripts/PDF%20and%20Blob%20Storage%20Example/PDFCreateandUploadtoAzureBlob.py
# Version: 0.0.7
# License: AGPL-3.0
# Author: Ben Swaby at FBCHville
# Editable: True

# ========== CONFIGURATION ==========
# Variables you can customize:

AZURE_ACCOUNT_KEY = ""  # Your account key
AZURE_ACCOUNT_NAME = ""  # Your storage account name
AZURE_CONTAINER_NAME = ""  # Your container name
PDF_FILENAME_PREFIX = "SomeFileName"
myVariable = 'test3222'
myVariable2 = 'abc123'

# =========  

# PCKGD_MANAGED_SECTION
# Do not edit below this line - managed by package updates

"""
=================================================================================
PDF and BLOB Storage EXAMPLE TEMPLATE
Created By: Ben Swaby
Email: bswaby@fbchtn.org 
=================================================================================
WHAT THIS IS:
This is an example template showing how to generate PDF from HTML and upload 
them to Azure Blob Storage within TouchPoint's IronPython environment. 
This is NOT production-ready code - you'll need to adapt it to
work with your specific TouchPoint data sources and workflow.

WHAT THIS DEMONSTRATES:
- How to convert HTML to PDF using iTextSharp
- How to upload PDFs to Azure Blob Storage
- How to get a shareable link to the PDF

USE CASE:
This example creates an IRA Qualified Charitable Distribution (QCD) letter,
but the same approach works for other HTML you need to generate 
and share.

PREREQUISITES:
1. TouchPoint CHMS with Python scripting enabled
2. iTextSharp and itextsharp.xmlworker libraries (check using code below)
3. Azure Blob Storage account (setup instructions below are best effort 
   of what I think needs done)

HOW TO USE THIS TEMPLATE:
1. Set up Azure Blob Storage (see instructions below)
2. Update the CONFIGURATION section with your Azure credentials
3. Replace the example HTML template with your letter format
4. Integrate with your TouchPoint data (replace placeholder variables)
5. Test with sample data before using in production

=================================================================================
AZURE BLOB STORAGE SETUP INSTRUCTIONS
=================================================================================

STEP 1: Create an Azure Storage Account
1. Go to https://portal.azure.com
2. Sign in or create a free Azure account (note: Check TechSoup for non-profit credits)
3. Click "Create a resource" > "Storage" > "Storage account"
4. Fill in the basics:
   - Subscription: Choose your subscription
   - Resource group: Create new or use existing
   - Storage account name: Choose a unique name (e.g., "yourchurchname")
   - Region: Choose closest to you
   - Performance: Standard (cheaper, sufficient for PDFs)
   - Redundancy: LRS (Locally-redundant storage) is fine for documents
5. Click "Review + Create" then "Create"
6. Wait for deployment to complete (1-2 minutes)

STEP 2: Create a Container
1. Go to your new storage account
2. In the left menu, click "Containers" under "Data storage"
3. Click "+ Container" at the top
4. Name it (e.g., "contribution-letters" or "documents")
5. Set "Public access level" to "Blob" if you want direct links to work
   (or keep it Private and generate SAS tokens for each file - more secure)
6. Click "Create"

STEP 3: Get Your Account Key
1. In your storage account, go to "Access keys" in the left menu
2. Click "Show keys" 
3. Copy the "Storage account name" - this is your AZURE_ACCOUNT_NAME
4. Copy "Key" from key1 or key2 - this is your AZURE_ACCOUNT_KEY
5. Your container name from Step 2 is your AZURE_CONTAINER_NAME

SECURITY NOTE: 
Your account key gives full access to your storage. Keep it secure! Don't
share it or commit it to version control. Consider using Azure Key Vault or
environment variables in production.

COST:
Azure Blob Storage is very inexpensive:
- Storage: ~$0.02 per GB per month
- Transactions: ~$0.05 per 10,000 operations
- Example: 1,000 PDFs/month (5MB each) = ~$0.10/month

FREE TIER:
New Azure accounts get $200 credit for 30 days, then Free Tier includes:
- 5 GB blob storage free for 12 months

=================================================================================
"""

# =================================================================================
# CONFIGURATION - UPDATE THESE VALUES FOR YOUR ENVIRONMENT
# =================================================================================

# Azure Storage Account Settings
# Replace these with your actual Azure credentials from setup instructions above
AZURE_ACCOUNT_NAME = ""  # Your storage account name
AZURE_ACCOUNT_KEY = ""  # Your account key
AZURE_CONTAINER_NAME = ""  # Your container name

# PDF File Naming Convention
# The PDF will be named: SomeFileName-YYYYMMDD-HHMMSS.pdf
PDF_FILENAME_PREFIX = "SomeFileName"

# =================================================================================
# EXAMPLE DATA - REPLACE WITH YOUR TOUCHPOINT DATA SOURCE
# =================================================================================
# In production, you would get these values from TouchPoint's Data object,
# person records, contribution records, etc. These are just placeholders to
# demonstrate the concept.
#
# For now, these are hardcoded examples:

DONOR_NAME = "John and Mary Smith"
GIFT_DATE = "December 15, 2024"
GIFT_AMOUNT = "5,000.00"
TAX_YEAR = "2024"

# =================================================================================
# HTML TEMPLATE - CUSTOMIZE FOR YOUR LETTER FORMAT
# =================================================================================
# IMPORTANT NOTES:
# - Use {donor_name}, {gift_date}, {gift_amount}, {year} as placeholders
# - Only use INLINE styles (style="...") - NO <style> blocks or <head> sections
# - Keep HTML simple - complex CSS may not render correctly
# - Test your HTML formatting before using in production
# - You can add more placeholder variables as needed
#
# Supported HTML tags: <p>, <strong>, <br/>, <ul>, <ol>, <li>, <table>, etc.
# Supported inline styles: text-align, font-weight, font-size, color, etc.

HTML_TEMPLATE = """
<p style="text-align: center;"><strong>IMPORTANT TAX DOCUMENTATION<br/>
RE: {year} IRA Qualified Charitable Distribution (QCD)</strong></p>

<p>Dear {donor_name},</p>

<p>Thank you for your charitable gift in the amount of ${gift_amount} from your Individual Retirement Account.  
We are writing to acknowledge that we received your gift on {gift_date} directly from your plan trustee/administrator/custodian.  
Therefore, all or a portion of your gift may qualify as a {year} qualified charitable distribution from your IRA under section 
408(d)(8) of the Internal Revenue Code and the Protecting Americans from Tax Hikes Act of 2015.</p>

<p>In that connection, we warrant to you that our organization is qualified as a public charity under section 170(b)(1)(A) 
of the Internal Revenue Code and that your gift was not transferred to either a donor advised fund or a supporting 
organization as described in section 509(a)(3).  We further warrant that no goods or services of any monetary value were 
or will be transferred to you in connection with this gift other than intangible religious benefits.  <strong>Please note:  
A QCD is not a tax-deductible charitable gift.  A QCD may, however, count towards the annual IRA required minimum 
distribution (RMD) and not be deemed taxable income.  Please consult with your own professional tax advisor regarding this 
and all appropriate matters.</strong>  Please retain this letter with your important tax documents and provide a copy to 
your tax preparer.</p>

<p>Thank you for your generous contribution in support of the ministries of First Baptist Church.  Together, we are sharing 
the love of the living Christ with a world in need.</p>

<p><br/>Cordially,<br/>[Signature]</p>

<p>Finance Minister</p>
"""

# =================================================================================
# MAIN CODE - You shouldn't need to modify this section
# =================================================================================
# This section handles the PDF generation and Azure upload process.
# Only modify if you need to change the core functionality.

import clr
clr.AddReference('iTextSharp')
clr.AddReference('itextsharp.xmlworker')

from iTextSharp.text import Document, PageSize
from iTextSharp.text.pdf import PdfWriter
from iTextSharp.text.html.simpleparser import HTMLWorker
from System.IO import MemoryStream, StringReader

import urllib
import urllib2
import base64
import datetime
import hmac
import hashlib

print("="*80)
print("IRA QCD Contribution Letter Generator - Example")
print("="*80)
print("This is a demonstration of PDF generation and Azure upload")
print("="*80)

# -----------------------------------------------------------------------------
# STEP 1: Prepare HTML content by replacing placeholders with actual values
# -----------------------------------------------------------------------------
print("\n[STEP 1] Preparing HTML content...")
print("Replacing placeholders with donor information...")

html_content = HTML_TEMPLATE.format(
    donor_name=DONOR_NAME,
    gift_date=GIFT_DATE,
    gift_amount=GIFT_AMOUNT,
    year=TAX_YEAR
)
print("  - Donor: {}".format(DONOR_NAME))
print("  - Gift Amount: ${}".format(GIFT_AMOUNT))
print("  - Gift Date: {}".format(GIFT_DATE))
print("HTML content prepared successfully")

# -----------------------------------------------------------------------------
# STEP 2: Generate PDF from HTML using iTextSharp
# -----------------------------------------------------------------------------
print("\n[STEP 2] Generating PDF from HTML...")
print("Using iTextSharp library to convert HTML to PDF...")

# Create a new PDF document in memory
memStream = MemoryStream()
document = Document(PageSize.LETTER)
document.SetMargins(50, 50, 50, 50)  # Set 50pt margins on all sides

# Get a PdfWriter instance to write to the memory stream
writer = PdfWriter.GetInstance(document, memStream)
document.Open()

# Parse the HTML and add it to the PDF document
stringReader = StringReader(html_content)
htmlWorker = HTMLWorker(document)
htmlWorker.Parse(stringReader)

document.Close()

# Convert the .NET MemoryStream to a byte array
pdfBytes = memStream.ToArray()
memStream.Close()

# Convert .NET byte array to Python bytes for HTTP upload
pdf_data = bytes(bytearray(pdfBytes))

print("PDF generated successfully!")
print("  - Size: {} bytes ({:.2f} KB)".format(len(pdf_data), len(pdf_data) / 1024.0))

# -----------------------------------------------------------------------------
# STEP 3: Prepare Azure Blob Storage upload
# -----------------------------------------------------------------------------
print("\n[STEP 3] Preparing Azure Blob Storage upload...")

# Generate a unique filename with timestamp to avoid overwrites
timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
blob_name = "{}-{}.pdf".format(PDF_FILENAME_PREFIX, timestamp)

# Construct the full URL where the blob will be stored
blob_url = "https://{}.blob.core.windows.net/{}/{}".format(
    AZURE_ACCOUNT_NAME, 
    AZURE_CONTAINER_NAME, 
    blob_name
)

print("  - Target URL: {}".format(blob_url))
print("  - Filename: {}".format(blob_name))

# Prepare HTTP PUT request details
method = "PUT"
now = datetime.datetime.utcnow()
date_string = now.strftime("%a, %d %b %Y %H:%M:%S GMT")

# -----------------------------------------------------------------------------
# STEP 4: Create Azure authorization signature
# -----------------------------------------------------------------------------
print("\n[STEP 4] Creating Azure authorization signature...")
print("Using HMAC-SHA256 to sign the request...")

# Azure Blob Storage requires a specific string format for signing
# See: https://docs.microsoft.com/en-us/rest/api/storageservices/authorize-with-shared-key
string_to_sign = (
    "{}\n"           # HTTP Verb (PUT)
    "\n"             # Content-Encoding (empty)
    "\n"             # Content-Language (empty)
    "{}\n"           # Content-Length
    "\n"             # Content-MD5 (empty)
    "application/pdf\n"  # Content-Type
    "\n"             # Date (empty, using x-ms-date instead)
    "\n"             # If-Modified-Since (empty)
    "\n"             # If-Match (empty)
    "\n"             # If-None-Match (empty)
    "\n"             # If-Unmodified-Since (empty)
    "\n"             # Range (empty)
    "x-ms-blob-type:BlockBlob\n"  # Custom header
    "x-ms-date:{}\n"               # Custom header
    "x-ms-version:2021-12-02\n"    # API version
    "/{}/{}/{}"      # CanonicalizedResource
).format(
    method, 
    len(pdf_data), 
    date_string, 
    AZURE_ACCOUNT_NAME, 
    AZURE_CONTAINER_NAME, 
    blob_name
)

# Sign the string using your account key
decoded_key = base64.b64decode(AZURE_ACCOUNT_KEY)
signature = base64.b64encode(
    hmac.new(decoded_key, string_to_sign.encode("utf-8"), hashlib.sha256).digest()
)
authorization_header = "SharedKey {}:{}".format(AZURE_ACCOUNT_NAME, signature)

print("Authorization signature created")

# -----------------------------------------------------------------------------
# STEP 5: Set up HTTP headers for Azure Blob Storage API
# -----------------------------------------------------------------------------
print("\n[STEP 5] Setting up HTTP headers...")

headers = {
    "x-ms-blob-type": "BlockBlob",        # Type of blob to create
    "x-ms-date": date_string,              # Request timestamp
    "x-ms-version": "2021-12-02",          # Azure API version
    "Content-Type": "application/pdf",     # MIME type
    "Content-Length": str(len(pdf_data)),  # File size
    "Authorization": authorization_header, # Signed authorization
}

print("Headers configured:")
for key in headers.keys():
    if key != "Authorization":  # Don't print the full auth header
        print("  - {}: {}".format(key, headers[key]))

# -----------------------------------------------------------------------------
# STEP 6: Create HTTP PUT request
# -----------------------------------------------------------------------------
print("\n[STEP 6] Creating HTTP PUT request...")

# IronPython 2.7's urllib2.Request doesn't support PUT by default
# We need to create a custom class that overrides get_method()
class PutRequest(urllib2.Request):
    def get_method(self):
        return "PUT"

# Create the PUT request with the PDF data
request = PutRequest(blob_url, data=pdf_data)

# Add all the required headers
for key, value in headers.items():
    request.add_header(key, value)

print("PUT request created with {} bytes of data".format(len(pdf_data)))

# -----------------------------------------------------------------------------
# STEP 7: Send request to Azure and handle response
# -----------------------------------------------------------------------------
print("\n[STEP 7] Uploading to Azure Blob Storage...")
print("Sending request... (this may take a few seconds)")

try:
    # Send the request
    response = urllib2.urlopen(request)
    
    # Check if upload was successful
    if response.getcode() == 201:  # 201 Created = success
        print("\n" + "="*80)
        print("SUCCESS! PDF UPLOADED SUCCESSFULLY")
        print("="*80)
        print("\nYour PDF has been uploaded to Azure Blob Storage")
        print("\n** DIRECT LINK TO PDF (copy and share this) **")
        print(blob_url)
        print("\n" + "="*80)
        print("\nNEXT STEPS:")
        print("1. Click the link above to verify the PDF looks correct")
        print("2. Use this link in TouchPoint emails (it won't be an attachment)")
        print("3. Recipients can click to view/download the PDF")
        print("4. Integrate this code with your TouchPoint workflow")
        print("="*80)
    else:
        print("\nWARNING: Unexpected response code: {}".format(response.getcode()))
        print("Response: {}".format(response.read()))
        
except urllib2.HTTPError as e:
    print("\n" + "="*80)
    print("ERROR: Upload failed")
    print("="*80)
    print("HTTP Error Code: {}".format(e.code))
    print("Error Details: {}".format(e.read()))
    print("\n** TROUBLESHOOTING TIPS **")
    print("1. Verify your AZURE_ACCOUNT_NAME is correct: '{}'".format(AZURE_ACCOUNT_NAME))
    print("2. Verify your AZURE_ACCOUNT_KEY is correct (check for copy/paste errors)")
    print("3. Ensure the container '{}' exists in your storage account".format(AZURE_CONTAINER_NAME))
    print("4. Check that your container's access level allows blob access")
    print("5. Verify your account key has write permissions")
    print("6. Make sure there are no firewall rules blocking the request")
    print("\nTo check your Azure setup:")
    print("- Go to https://portal.azure.com")
    print("- Navigate to your storage account: {}".format(AZURE_ACCOUNT_NAME))
    print("- Check 'Containers' to verify '{}' exists".format(AZURE_CONTAINER_NAME))
    print("- Check 'Access keys' to verify your key is correct")
    print("="*80)
    
except Exception as e:
    print("\n" + "="*80)
    print("ERROR: Unexpected error occurred")
    print("="*80)
    print("Error Type: {}".format(type(e).__name__))
    print("Error Message: {}".format(str(e)))
    print("\nThis might be a network issue or a problem with the IronPython environment.")
    print("="*80)

print("\n" + "="*80)
print("Process complete")
print("="*80)
