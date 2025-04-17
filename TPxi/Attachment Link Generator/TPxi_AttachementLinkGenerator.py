"""
TouchPoint Attachment Link Generator

This script generates direct download links for attachments in TouchPoint's OrgMemberDocuments table.
Unlike automated downloaders, this approach uses the browser's native authenticated session to access the files.

Features:
- View all attachments for an organization
- Get direct download links for each document
- Display document information in a clean interface

--Upload Instructions Start--
To upload code to TouchPoint, use the following steps.
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python script (e.g., "AttachmentLinks") and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--
"""

# Import necessary modules
import re
import traceback

# Set script title
model.Title = "Attachment Link Generator"

# Batch size for displaying attachments per page
BATCH_SIZE = 20


# TouchPoint Domain - You shouldn't need to change this
TOUCHPOINT_DOMAIN = model.CmsHost  


def sanitize_filename(filename):
    """Sanitize a filename"""
    # Ensure input is a string
    filename = str(filename)
    
    # Replace special characters
    filename = re.sub(r'[\\/*?:"<>|]', '-', filename)
    filename = re.sub(r'-+', '-', filename)
    filename = filename.strip('-')
    return filename

def get_org_info(org_id):
    """Get organization info"""
    try:
        sql = "SELECT OrganizationId, OrganizationName FROM Organizations WHERE OrganizationId = " + str(org_id)
        result = q.QuerySqlTop1(sql)
        if result:
            return {"id": result.OrganizationId, "name": result.OrganizationName}
        return None
    except Exception as e:
        print "<div class='error'>Error getting organization info: " + str(e) + "</div>"
        return None

def get_org_attachments(org_id):
    """Get attachments for an organization"""
    try:
        sql = """
        SELECT 
            o.OrganizationId,
            o.OrganizationName,
            doc.DocumentId,
            doc.DocumentName,
            doc.PeopleId,
            p.PeopleId,
            p.Name,
            p.Name2,
            doc.CreatedDate
        FROM OrgMemberDocuments doc
        JOIN Organizations o ON doc.OrganizationId = o.OrganizationId
        JOIN People p ON doc.PeopleId = p.PeopleId
        WHERE doc.OrganizationId = """ + str(org_id) + """
        ORDER BY doc.CreatedDate DESC
        """
        return q.QuerySql(sql)
    except Exception as e:
        print "<div class='error'>Error querying attachments: " + str(e) + "</div>"
        return None

def get_document_url(document_id):
    """Generate download URL"""
    return TOUCHPOINT_DOMAIN + "/OrgMemberDialog/DocumentDownload/" + str(document_id)

# CSS styles
styles = """
<style>
    .container {
        max-width: 1000px;
        margin: 0 auto;
        padding: 20px;
        font-family: Arial, sans-serif;
    }
    .form-group {
        margin-bottom: 15px;
    }
    label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
    }
    input[type="text"] {
        width: 100%;
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
    }
    button, .btn {
        background-color: #4CAF50;
        color: white;
        padding: 10px 15px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        margin-right: 5px;
        text-decoration: none;
        display: inline-block;
    }
    button:hover, .btn:hover {
        background-color: #45a049;
    }
    .btn-secondary {
        background-color: #3498db;
    }
    .btn-secondary:hover {
        background-color: #2980b9;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
    }
    table th, table td {
        padding: 8px;
        text-align: left;
        border-bottom: 1px solid #ddd;
    }
    table th {
        background-color: #f2f2f2;
    }
    .loading {
        display: none;
        text-align: center;
        margin: 20px 0;
    }
    .spinner {
        border: 4px solid #f3f3f3;
        border-top: 4px solid #3498db;
        border-radius: 50%;
        width: 30px;
        height: 30px;
        animation: spin 2s linear infinite;
        margin: 0 auto;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    .success {
        color: green;
        font-weight: bold;
    }
    .error {
        color: red;
        font-weight: bold;
    }
    .alert {
        padding: 15px;
        margin-bottom: 20px;
        border: 1px solid transparent;
        border-radius: 4px;
    }
    .alert-info {
        color: #31708f;
        background-color: #d9edf7;
        border-color: #bce8f1;
    }
    .alert-success {
        color: #3c763d;
        background-color: #dff0d8;
        border-color: #d6e9c6;
    }
    .pagination {
        margin-top: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .download-link {
        color: #3498db;
        text-decoration: none;
        display: inline-block;
        margin-right: 10px;
    }
    .download-link:hover {
        text-decoration: underline;
    }
    .action-cell {
        display: flex;
        gap: 5px;
    }
    /* Add a tooltip style */
    .tooltip {
        position: relative;
        display: inline-block;
    }
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 250px;
        background-color: #555;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 5px;
        position: absolute;
        z-index: 1;
        bottom: 125%;
        left: 50%;
        margin-left: -125px;
        opacity: 0;
        transition: opacity 0.3s;
    }
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
    /* Copy button style */
    .copy-button {
        background-color: #ddd;
        border: none;
        color: #333;
        padding: 5px 10px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 12px;
        border-radius: 4px;
        cursor: pointer;
    }
    .copy-button:hover {
        background-color: #ccc;
    }
    /* Download instructions */
    .download-instructions {
        padding: 15px;
        margin-top: 20px;
        background-color: #f9f9f9;
        border-left: 5px solid #3498db;
    }
    .download-instructions h3 {
        margin-top: 0;
        color: #3498db;
    }
    .download-instructions ol {
        margin-left: 20px;
    }
</style>
"""

# JavaScript
scripts = """
<script>
    function showLoading() {
        document.getElementById('loading').style.display = 'block';
    }
    
    // Function to copy URL to clipboard
    function copyToClipboard(text, buttonId) {
        // Create temporary element
        var tempInput = document.createElement("input");
        tempInput.value = text;
        document.body.appendChild(tempInput);
        
        // Select and copy
        tempInput.select();
        document.execCommand("copy");
        
        // Remove temporary element
        document.body.removeChild(tempInput);
        
        // Update button text
        var button = document.getElementById(buttonId);
        var originalText = button.innerHTML;
        button.innerHTML = "Copied!";
        
        // Revert button text after 2 seconds
        setTimeout(function() {
            button.innerHTML = originalText;
        }, 2000);
    }
    
    // Function to open multiple URLs in new tabs
    function openSelectedLinks() {
        var checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
        var delay = 300; // Delay in milliseconds between opening tabs
        
        if (checkboxes.length > 10) {
            var proceed = confirm("You're about to open " + checkboxes.length + " tabs. This may slow down your browser. Continue?");
            if (!proceed) return;
        }
        
        checkboxes.forEach(function(checkbox, index) {
            setTimeout(function() {
                var url = checkbox.getAttribute('data-url');
                if (url) {
                    window.open(url, '_blank');
                }
            }, index * delay);
        });
    }
    
    // Function to toggle all checkboxes
    function toggleAll(source) {
        var checkboxes = document.querySelectorAll('input[type="checkbox"][name^="doc_cb_"]');
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checked = source.checked;
        }
    }
</script>
"""

# Main script
try:
    # Print CSS
    print styles
    
    # Get parameters
    org_id = model.Data.org_id
    start_index = 0
    
    try:
        start_index = int(model.Data.start) if model.Data.start else 0
    except:
        start_index = 0
    
    # Show base container
    print "<div class='container'>"
    print """<h1>Attachment Link Generator
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
                    <!-- Text portion - TP -->
                    <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
                    
                    <!-- Circular element -->
                    <g transform="translate(190, 107)">
                      <!-- Outer circle -->
                      <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
                      
                      <!-- White middle circle -->
                      <circle cx="0" cy="0" r="10.5" fill="white"/>
                      
                      <!-- Inner circle -->
                      <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
                      
                      <!-- X crossing through the circles -->
                      <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
                    </g>
                    
                    <!-- Single "i" letter to the right -->
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
                  </svg>
             </h1>"""
    
    # Show organization form
    print """
    <form id='orgForm' method='get' action=''>
        <div class='form-group'>
            <label for='org_id'>Involvement ID:</label>
            <input type='text' id='org_id' name='org_id' value='""" + (str(org_id) if org_id else "") + """' required>
        </div>
        <button type='submit' onclick='showLoading()'>Load Attachments</button>
    </form>
    """
    
    print "<div id='loading' class='loading'>"
    print "<div class='spinner'></div>"
    print "<p>Loading attachments...</p>"
    print "</div>"
    
    if org_id:
        # Get organization info
        org_info = get_org_info(org_id)
        
        if not org_info:
            print "<div class='error'>Error: Organization ID " + str(org_id) + " not found.</div>"
        else:
            # Get attachments
            attachments = get_org_attachments(org_id)
            
            print "<h2>Attachments for Involvement: " + org_info["name"] + "</h2>"
            
            if attachments is None:
                print "<div class='error'>Error loading attachments.</div>"
            elif len(attachments) == 0:
                print "<p>No attachments found for this organization.</p>"
            else:
                # Add instructions for downloading
                print """
                <div class="download-instructions">
                    <h3>How to Download Files</h3>
                    <p>Since TouchPoint requires authentication to download files, this tool provides direct links to each document.</p>
                    <ol>
                        <li>Use the "Download" button to download files one at a time, or</li>
                        <li>Check the boxes next to documents you want to download, then click "Open Selected in New Tabs"</li>
                        <li>If your browser blocks the pop-ups, you'll need to allow them</li>
                        <li>Save each file from its tab, renaming it if needed</li>
                    </ol>
                </div>
                """
                
                # Add button to open selected documents in new tabs
                print """
                <div style="margin: 20px 0;">
                    <button type="button" onclick="openSelectedLinks()" class="btn-secondary">Open Selected in New Tabs</button>
                </div>
                """
                
                print "<table>"
                print "<thead>"
                print "<tr>"
                print "<th><input type='checkbox' id='selectAll' onclick='toggleAll(this)' checked></th>"
                print "<th>Person</th>"
                print "<th>Document Name</th>"
                print "<th>Created Date</th>"
                print "<th>Actions</th>"
                print "</tr>"
                print "</thead>"
                print "<tbody>"
                
                # Convert to list first
                attachment_list = []
                for a in attachments:
                    attachment_list.append(a)
                
                # Display limited attachments
                batch_limit = BATCH_SIZE
                end_index = min(len(attachment_list), start_index + batch_limit)
                
                for i in range(start_index, end_index):
                    # Create individual download link
                    doc_url = get_document_url(attachment_list[i].DocumentId)
                    doc_name = attachment_list[i].DocumentName
                    
                    # Add .jpg extension if not present (since that's how TouchPoint serves it)
                    suggested_filename = doc_name
                    if not suggested_filename.lower().endswith('.jpg'):
                        suggested_filename += '.jpg'
                    
                    # Generate a unique ID for this button
                    button_id = "copy_btn_" + str(attachment_list[i].DocumentId)
                    
                    print "<tr>"
                    print "<td><input type='checkbox' name='doc_cb_" + str(i) + "' data-url='" + doc_url + "' checked></td>"
                    print "<td>" + str(attachment_list[i].Name) + " (" + str(attachment_list[i].PeopleId) + ")</td>"
                    print "<td>" + str(doc_name) + "</td>"
                    print "<td>" + str(attachment_list[i].CreatedDate) + "</td>"
                    print "<td class='action-cell'>"
                    
                    # Direct download link
                    print "<a href='" + doc_url + "' class='download-link' download='" + suggested_filename + "' target='_blank'>Download</a>"
                    
                    # Copy URL button with tooltip
                    print "<div class='tooltip'>"
                    print "<button id='" + button_id + "' class='copy-button' onclick=\"copyToClipboard('" + doc_url + "', '" + button_id + "')\">Copy URL</button>"
                    print "<span class='tooltiptext'>Copy link to clipboard</span>"
                    print "</div>"
                    
                    print "</td>"
                    print "</tr>"
                
                print "</tbody>"
                print "</table>"
                
                # Pagination controls
                if len(attachment_list) > batch_limit:
                    print "<div class='pagination'>"
                    if start_index > 0:
                        prev_start = max(0, start_index - batch_limit)
                        print "<a href='?org_id=" + str(org_id) + "&start=" + str(prev_start) + "' class='btn'>Previous</a>"
                    
                    if end_index < len(attachment_list):
                        next_start = end_index
                        print "<a href='?org_id=" + str(org_id) + "&start=" + str(next_start) + "' class='btn'>Next</a>"
                    
                    print "<span>Showing " + str(start_index + 1) + " to " + str(end_index) + " of " + str(len(attachment_list)) + " attachments</span>"
                    print "</div>"
    
    # Close container
    print "</div>"
    
    # Add JavaScript
    print scripts
    
except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
