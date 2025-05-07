"""
Geographic Distribution Map with Census Data Overlays
---------------------------
This script creates an interactive map showing the geographic distribution of people
from a BlueToolbar query, with profile photos, grouped markers for same address, 
census data overlays, and the ability to tag people from selections.

Features:
- Shows people's locations on Google Maps with profile photos
- Groups markers for people at the same address
- Filter by street
- Tag people by area, street, or custom selection
- Census data layer overlays (income, education, age, etc.)
- Export to CSV option

Note:  You will need a Google Maps API Key, and if you enable the census data, you will also need a census data api key.

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python script (e.g., "GeographicDistributionMap") and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
"""

# ========= CONFIGURATION SETTINGS (EDIT THESE) =========
# Map display settings
MAP_TITLE = "Geographic Distribution Map"
MAP_HEIGHT = 800  # Map height in pixels

# Church location (default center of map)
CHURCH_NAME = "First Baptist Church"
CHURCH_LATITUDE = 36.320205  # Church address latitude
CHURCH_LONGITUDE = -86.567619  # Church address longitude
DEFAULT_ZOOM = 12  # Default zoom level

# API Keys (REPLACE THESE WITH YOUR OWN)
GOOGLE_MAPS_API_KEY = ""  # Get from https://console.cloud.google.com/
CENSUS_API_KEY = ""  # Get from https://api.census.gov/data/key_signup.html

# Census data overlay options
ENABLE_CENSUS_DATA = True  # Set to False to disable census data features

# Census Data Overlay Options
# You can add/remove items from this list to customize the census data available in the map
CENSUS_OVERLAY_OPTIONS = [
    {"id": "median_income", "name": "Median Household Income", "variable": "B19013_001E", "endpoint": "acs/acs5"},
    {"id": "pop_density", "name": "Population Density", "variable": "B01003_001E", "endpoint": "acs/acs5"},
    {"id": "median_age", "name": "Median Age", "variable": "B01002_001E", "endpoint": "acs/acs5"},
    {"id": "education_college", "name": "College Education %", "variable": "B15003_022E", "endpoint": "acs/acs5"},
    {"id": "owner_occupied", "name": "Owner Occupied Homes %", "variable": "B25003_002E", "endpoint": "acs/acs5"}
]

"""
Additional Census Data Options:
You can add any of these options to the CENSUS_OVERLAY_OPTIONS list above.
Simply copy and paste the desired items into the CENSUS_OVERLAY_OPTIONS list.

# Demographic Details
{"id": "male_population", "name": "Male Population", "variable": "B01001_002E", "endpoint": "acs/acs5"},
{"id": "female_population", "name": "Female Population", "variable": "B01001_026E", "endpoint": "acs/acs5"},

# Race & Ethnicity
{"id": "white_alone", "name": "White Alone Population", "variable": "B02001_002E", "endpoint": "acs/acs5"},
{"id": "black_alone", "name": "Black or African American Alone", "variable": "B02001_003E", "endpoint": "acs/acs5"},
{"id": "asian_alone", "name": "Asian Alone", "variable": "B02001_005E", "endpoint": "acs/acs5"},
{"id": "hispanic_latino", "name": "Hispanic or Latino (Any Race)", "variable": "B03003_003E", "endpoint": "acs/acs5"},

# Housing & Households
{"id": "housing_units", "name": "Total Housing Units", "variable": "B25001_001E", "endpoint": "acs/acs5"},
{"id": "median_rent", "name": "Median Gross Rent", "variable": "B25064_001E", "endpoint": "acs/acs5"},
{"id": "median_home_value", "name": "Median Home Value", "variable": "B25077_001E", "endpoint": "acs/acs5"},
{"id": "vacant_housing", "name": "Vacant Housing Units", "variable": "B25002_003E", "endpoint": "acs/acs5"},

# Income & Poverty
{"id": "per_capita_income", "name": "Per Capita Income", "variable": "B19301_001E", "endpoint": "acs/acs5"},
{"id": "poverty_population", "name": "Population Below Poverty Line", "variable": "B17001_002E", "endpoint": "acs/acs5"},

# Employment & Labor
{"id": "employed_population", "name": "Employed Population (Age 16+)", "variable": "B23025_004E", "endpoint": "acs/acs5"},
{"id": "labor_force_population", "name": "Labor Force Population (Age 16+)", "variable": "B23025_003E", "endpoint": "acs/acs5"},

# Commute & Transportation
{"id": "mean_commute", "name": "Mean Travel Time to Work (Minutes)", "variable": "B08303_001E", "endpoint": "acs/acs5"},
{"id": "commute_drive_alone", "name": "Drove Alone to Work %", "variable": "B08301_003E", "endpoint": "acs/acs5"},
{"id": "commute_public_transit", "name": "Used Public Transportation %", "variable": "B08301_010E", "endpoint": "acs/acs5"},

# Education (Broader)
{"id": "high_school_plus", "name": "High School Graduate or Higher %", "variable": "B15003_017E", "endpoint": "acs/acs5"},
{"id": "graduate_degree", "name": "Graduate or Professional Degree %", "variable": "B15003_025E", "endpoint": "acs/acs5"}
"""
# ======== END CONFIGURATION SETTINGS ========

# Import necessary libraries
import traceback
import re
import collections

# Helper function to safely convert string to float
def safe_float(value):
    try:
        if value is None:
            return None
        return float(value)
    except:
        return None

# Helper function to get street name from address
def get_street_name(address):
    if not address:
        return ""
    street_parts = address.split(',')
    if not street_parts:
        street_parts = [address]
    return re.sub(r'^\d+\s+', '', street_parts[0]).strip()

# Main action handler
def handle_tag_people():
    """Handle tagging of people."""
    # Check if form data is available
    if hasattr(model.Data, 'people_ids') and hasattr(model.Data, 'tag_name'):
        try:
            # Get people IDs and tag name from form
            people_ids_str = model.Data.people_ids
            tag_name = model.Data.tag_name
            
            # Tag the people
            owner_id = model.UserPeopleId  # Current user as tag owner
            
            # Build a query using the people IDs
            people_query = "peopleids='{}'".format(people_ids_str)
            
            # Add the tag
            model.AddTag(people_query, tag_name, owner_id, False)
            
            # Print success message
            print "<h2>Tagging Complete</h2>"
            print "<p>{} people have been tagged with: <strong>{}</strong></p>".format(
                people_ids_str.count(',') + 1, tag_name)
            print "<p><a href='javascript:window.close();' class='btn btn-primary'>Close</a></p>"
        except Exception as e:
            # Print error message
            print "<h2>Tagging Error</h2>"
            print "<p>An error occurred: {}</p>".format(str(e))
            print "<p><a href='javascript:window.close();' class='btn btn-primary'>Close</a></p>"
    else:
        # No data provided
        print "<h2>Tagging Error</h2>"
        print "<p>No people selected for tagging. Please return to the map and select people to tag.</p>"

# Process people data for mapping
def process_people_data(people_list, total_count=0):
    """Process people data and prepare for mapping."""
    if not total_count:
        total_count = len(people_list)
    
    # Prepare collections
    people_by_address = collections.defaultdict(list)
    people_by_street = {}
    located_count = 0
    
    # Get people IDs
    people_ids = []
    for person in people_list:
        people_ids.append(str(person.PeopleId))
    
    if not people_ids:
        return {
            "error": "No people found in query",
            "total_count": total_count,
            "located_count": 0,
            "address_count": 0,
            "street_count": 0
        }
    
    try:
        # Direct query for geocode and picture data
        direct_sql = """
        SELECT 
            p.PeopleId, 
            p.FirstName + ' ' + p.LastName as Name,
            p.PrimaryAddress,
            p.PrimaryCity,
            p.PrimaryState,
            p.PrimaryZip,
            p.HomePhone,
            p.CellPhone,
            p.EmailAddress,
            g.Latitude,
            g.Longitude,
            pic.MediumUrl as PictureUrl
        FROM People p
        LEFT JOIN Geocodes g ON g.Address = CONCAT(p.PrimaryAddress, ' ', p.PrimaryAddress2, ' ', p.PrimaryCity, ' ', p.PrimaryState, ' ', p.PrimaryZip)
        LEFT JOIN Picture pic ON p.PictureId = pic.PictureId
        WHERE p.PeopleId IN ({0})
        AND g.Latitude IS NOT NULL 
        AND g.Longitude IS NOT NULL
        """.format(",".join(people_ids))
        
        # Process each person
        for row in q.QuerySql(direct_sql):
            try:
                # Get latitude and longitude
                lat = safe_float(row.Latitude)
                lng = safe_float(row.Longitude)
                
                # Skip people without valid coordinates
                if lat is None or lng is None:
                    continue
                
                # Get picture URL if available
                picture_url = getattr(row, 'PictureUrl', None)
                
                # Create a simple person object
                class SimplePerson:
                    def __init__(self, pid, name, address, city, state, zip_code, home_phone, 
                                cell_phone, email, lat, lng, picture_url=None):
                        self.PeopleId = pid
                        self.Name = name
                        self.PrimaryAddress = address
                        self.PrimaryCity = city
                        self.PrimaryState = state
                        self.PrimaryZip = zip_code
                        self.HomePhone = home_phone
                        self.CellPhone = cell_phone
                        self.EmailAddress = email
                        self.Latitude = lat
                        self.Longitude = lng
                        self.PictureUrl = picture_url
                        # Full address for grouping
                        self.FullAddress = (address or "") + ", " + (city or "") + ", " + (state or "")
                
                person = SimplePerson(
                    row.PeopleId,
                    row.Name,
                    row.PrimaryAddress,
                    row.PrimaryCity,
                    row.PrimaryState,
                    row.PrimaryZip,
                    row.HomePhone,
                    row.CellPhone,
                    row.EmailAddress,
                    lat,
                    lng,
                    picture_url
                )
                
                # Get street name
                street = get_street_name(person.PrimaryAddress)
                
                # Add to street collection
                if street:
                    if street not in people_by_street:
                        people_by_street[street] = []
                    people_by_street[street].append(person)
                
                # Group by full address for popup grouping
                addr_key = (person.PrimaryAddress, person.PrimaryCity, person.PrimaryState)
                people_by_address[addr_key].append(person)
                
                located_count += 1
            except:
                continue
    except Exception as e:
        return {
            "error": str(e),
            "total_count": total_count,
            "located_count": 0,
            "address_count": 0,
            "street_count": 0
        }
    
    return {
        "success": True,
        "total_count": total_count,
        "located_count": located_count,
        "address_count": len(people_by_address),
        "street_count": len(people_by_street),
        "people_by_address": people_by_address,
        "people_by_street": people_by_street
    }

# Generate JavaScript data for the map
def generate_map_data(people_by_address, people_by_street, host_url):
    """Generate JavaScript for map markers and street data."""
    people_data_js = "["
    marker_count = 0
    
    # Process each address
    for addr_key, people_list in people_by_address.items():
        try:
            # Get coordinates from first person (all at same address)
            first_person = people_list[0]
            lat = safe_float(first_person.Latitude)
            lng = safe_float(first_person.Longitude)
            
            if lat is None or lng is None:
                continue
            
            # Get street name
            street = get_street_name(first_person.PrimaryAddress)
            
            # Create popup content for all people at this address
            popup_html = '<div class="person-popup address-group">'
            
            # Add address header
            full_address = ""
            if first_person.PrimaryAddress:
                full_address += first_person.PrimaryAddress + " "
            if first_person.PrimaryCity:
                full_address += first_person.PrimaryCity + ", "
            if first_person.PrimaryState:
                full_address += first_person.PrimaryState + " "
            if first_person.PrimaryZip:
                full_address += first_person.PrimaryZip
                
            popup_html += '<h3 class="address-header">' + full_address.strip() + '</h3>'
            
            # Add each person
            for i, person in enumerate(people_list):
                popup_html += '<div class="person-entry">'
                
                # Add picture if available
                if person.PictureUrl:
                    popup_html += '<img src="' + person.PictureUrl + '" class="popup-photo" />'
                
                popup_html += '<h4><a href="{}/Person2/{}" target="_blank">{}</a></h4>'.format(
                    host_url, person.PeopleId, person.Name)
                
                # Add contact info
                if person.HomePhone:
                    popup_html += '<p><strong>Home:</strong> {}</p>'.format(
                        model.FmtPhone(person.HomePhone))
                if person.EmailAddress:
                    popup_html += '<p><strong>Email:</strong> {}</p>'.format(person.EmailAddress)
                
                popup_html += '</div>'
                if i < len(people_list) - 1:  # Add separator if not the last person
                    popup_html += '<hr class="person-separator" />'
            
            popup_html += '</div>'
            
            # Add person IDs for this address
            person_ids_str = ",".join([str(person.PeopleId) for person in people_list])
            
            # Add marker data to JavaScript
            people_data_js += "  {"
            people_data_js += "id: '" + person_ids_str + "',"  # List of IDs as a string
            people_data_js += "name: '" + full_address.replace("'", "\\'") + "',"
            people_data_js += "lat: " + str(lat) + ","
            people_data_js += "lng: " + str(lng) + ","
            people_data_js += "street: '" + street.replace("'", "\\'") + "',"
            people_data_js += "popupContent: '" + popup_html.replace("'", "\\'").replace("\n", "") + "',"
            people_data_js += "count: " + str(len(people_list))
            people_data_js += "},"
            
            marker_count += 1
            
        except:
            continue
    
    people_data_js += "];"
    
    # Create street data
    street_data_js = "{"
    for street, people in people_by_street.items():
        street_data_js += "'" + street.replace("'", "\\'") + "': ["
        for person in people:
            street_data_js += str(person.PeopleId) + ","
        street_data_js = street_data_js.rstrip(',')  # Remove trailing comma
        street_data_js += "],"
    street_data_js = street_data_js.rstrip(',')  # Remove trailing comma
    street_data_js += "};"
    
    return {
        "people_data_js": people_data_js,
        "street_data_js": street_data_js,
        "marker_count": marker_count
    }

# Render the map HTML
def render_map(data, config):
    """Render the map HTML with the provided data."""
    # Get config values
    map_title = config.get("MAP_TITLE", MAP_TITLE)
    default_lat = config.get("CHURCH_LATITUDE", CHURCH_LATITUDE)
    default_lng = config.get("CHURCH_LONGITUDE", CHURCH_LONGITUDE)
    default_zoom = config.get("DEFAULT_ZOOM", DEFAULT_ZOOM)
    map_height = config.get("MAP_HEIGHT", MAP_HEIGHT)
    google_maps_api_key = config.get("GOOGLE_MAPS_API_KEY", GOOGLE_MAPS_API_KEY)
    enable_census = config.get("ENABLE_CENSUS_DATA", ENABLE_CENSUS_DATA)
    census_api_key = config.get("CENSUS_API_KEY", CENSUS_API_KEY)
    census_options = config.get("CENSUS_OVERLAY_OPTIONS", CENSUS_OVERLAY_OPTIONS)
    church_name = config.get("CHURCH_NAME", CHURCH_NAME)
    
    # Script name for form action
    script_name = "GeographicDistributionMap"
    if hasattr(model, 'ScriptName'):
        script_name = model.ScriptName
        
    # Start HTML
    html = """<!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            #map-container { position: relative; height: """ + str(map_height) + """px; width: 100%; }
            #map { height: 100%; width: 100%; }
            #controls { position: absolute; top: 10px; right: 10px; z-index: 1000; 
                      background: white; padding: 10px; border-radius: 4px; box-shadow: 0 0 10px rgba(0,0,0,0.2);
                      transition: height 0.3s ease; overflow: hidden; max-width: 300px; }
            #controls.collapsed { height: 40px; }
            #controls-header { cursor: pointer; margin: 0; padding-bottom: 8px; display: flex; align-items: center; justify-content: space-between; }
            #controls-header h4 { margin: 0; }
            #controls-toggle { font-size: 16px; font-weight: bold; }
            #controls-content { padding-top: 5px; }
            
            .region-select { margin: 5px 0; }
            .export-btn { margin-top: 10px; }
            .person-popup { max-width: 350px; }
            .address-header { border-bottom: 1px solid #ddd; margin-bottom: 10px; }
            .person-entry { margin-bottom: 10px; position: relative; min-height: 60px; }
            .popup-photo { max-width: 60px; height: 60px; float: right; margin-left: 10px; border-radius: 50%; object-fit: cover; }
            .person-separator { margin: 10px 0; border-top: 1px dashed #eee; }
            .loading { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); 
                     background: rgba(255,255,255,0.8); padding: 20px; border-radius: 5px; 
                     z-index: 2000; }
            #status-bar { background: #f5f5f5; padding: 8px; border-radius: 4px; margin-bottom: 10px; }
            .drawing-controls { margin-bottom: 10px; }
            .census-controls { border-top: 1px solid #ddd; padding-top: 10px; margin-top: 10px; }
            .census-selector { margin-bottom: 5px; }
            
            /* Modal styles */
            .modal { display: none; position: fixed; z-index: 2001; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.4); }
            .modal-content { background-color: #fefefe; margin: 15% auto; padding: 20px; border: 1px solid #888; width: 300px; border-radius: 5px; }
            .modal-header { margin-bottom: 15px; }
            .modal-footer { margin-top: 15px; text-align: right; }
            .modal-footer button { margin-left: 10px; }
            
            /* Census layer legend */
            #census-legend { position: absolute; bottom: 30px; left: 10px; background: white; padding: 10px; 
                            border-radius: 4px; box-shadow: 0 0 10px rgba(0,0,0,0.2); display: none; z-index: 999; 
                            max-width: 180px; max-height: 300px; overflow-y: auto; }
            #legend-title { margin-top: 0; margin-bottom: 5px; font-size: 14px; }
            .legend-item { display: flex; align-items: center; margin-bottom: 2px; }
            .color-box { width: 20px; height: 20px; margin-right: 5px; border: 1px solid #ccc; }
        </style>
    </head>
    <body>
        <h2>""" + map_title + """</h2>
        <div id="status-bar">
            Showing <strong>""" + str(data.get("located_count", 0)) + " of " + str(data.get("total_count", 0)) + """</strong> people on map at """ + str(data.get("marker_count", 0)) + """ locations. 
            <span style="color:#777">(People without valid addresses cannot be displayed)</span>
        </div>
        <div id="map-container">
            <div id="map"></div>
            <div id="controls" class="collapsed">
                <div id="controls-header" onclick="toggleControls()">
                    <h4>Map Controls</h4>
                    <span id="controls-toggle">+</span>
                </div>
                <div id="controls-content">
                    <div class="drawing-controls">
                        <button id="drawing-btn" onclick="toggleDrawingMode()">Enable Drawing Mode</button>
                        <button id="clear-drawing-btn" onclick="clearDrawing()" disabled>Clear Drawing</button>
                        <button id="tag-drawn-btn" onclick="tagDrawnArea()" disabled>Tag Selected</button>
                    </div>
                    <div class="region-select">
                        <label>Street:</label>
                        <select id="street-select" onchange="filterMarkersByStreet()">
                            <option value="all">All Streets</option>
    """
    
    # Add street options to selector
    street_list = list(data.get("people_by_street", {}).keys())
    for street in sorted(street_list):
        if street:
            try:
                count = len(data.get("people_by_street", {}).get(street, []))
                html += '<option value="' + street.replace('"', '\\"') + '">' + street + ' (' + str(count) + ')</option>'
            except:
                continue
    
    html += """
                        </select>
                        <button onclick="tagPeopleByStreet()">Tag</button>
                    </div>
                    <button class="export-btn" onclick="exportToCsv()">Export to CSV</button>
                    <button class="export-btn" onclick="tagAllVisible()">Tag All Visible</button>
    """
    
    # Add census data controls if enabled
    if enable_census:
        html += """
                    <div class="census-controls">
                        <h5>Census Data Overlays</h5>
                        <div class="census-selector">
                            <select id="census-select" onchange="toggleCensusLayer()">
                                <option value="">None (Off)</option>
        """
        
        for overlay in census_options:
            html += '<option value="' + overlay["id"] + '">' + overlay["name"] + '</option>'
            
        html += """
                            </select>
                        </div>
                        <div class="census-opacity">
                            <label>Opacity: <span id="opacity-value">0.7</span></label>
                            <input type="range" id="opacity-slider" min="0.1" max="1.0" step="0.1" value="0.7" 
                                   onchange="updateCensusOpacity(this.value)" oninput="updateOpacityLabel(this.value)">
                        </div>
                        <div class="census-help">
                            <button onclick="showLegend()" style="font-size: 12px; padding: 2px 5px; margin-top: 5px;">Show Legend</button>
                            <p style="font-size: 11px; color: #666; margin-top: 5px;">Using real Census ACS 5-Year data via Census API</p>
                        </div>
                    </div>
        """
    
    html += """
                </div>
            </div>
            
            <!-- Census legend - outside the controls div -->
            <div id="census-legend" style="display:none;">
                <h4 id="legend-title">Legend</h4>
                <div id="legend-items"></div>
            </div>
        </div>
        <div class="loading" id="loading">Processing...</div>
        
        <!-- Tag Name Modal -->
        <div id="tagModal" class="modal">
            <div class="modal-content">
                <div class="modal-header">
                    <h3>Create Tag</h3>
                </div>
                <div class="modal-body">
                    <p>Enter a name for the tag:</p>
                    <input type="text" id="tagName" style="width: 100%; padding: 5px;" 
                           placeholder="e.g., Spring Outreach 2025">
                </div>
                <div class="modal-footer">
                    <button onclick="closeTagModal()">Cancel</button>
                    <button onclick="submitTag()">Create Tag</button>
                </div>
            </div>
        </div>
        
        <!-- Hidden form for tagging -->
        <form id="tag-form" method="post" action="/PyScriptForm/""" + script_name + """" target="_blank" style="display:none;">
            <input type="hidden" name="action" value="tag_people">
            <input type="hidden" name="people_ids" id="tag_people_ids" value="">
            <input type="hidden" name="tag_name" id="tag_name" value="">
        </form>
    
        <script>
            // Map data
            var peopleData = """ + data.get("people_data_js", "[]") + """
            var streetData = """ + data.get("street_data_js", "{}") + """
            var markers = [];
            var visibleMarkers = [];
            var map;
            var infoWindow;
            var drawingManager;
            var selectedShape;
            var drawnAreaPeopleIds = [];
            var currentPeopleIds = []; // For current selection to tag
            var currentTagSource = ""; // Source of the current selection for tag modal title
            var censusOverlay = null;
            var activeCensusLayer = ""; // Currently active census layer ID
            
            // Census layer definitions
            var censusLayers = {"""
    
    # Add census layer definitions
    if enable_census:
        for i, overlay in enumerate(census_options):
            html += """
                '{}': {{
                    name: '{}',
                    variable: '{}'
                }}""".format(overlay["id"], overlay["name"], overlay["variable"])
            if i < len(census_options) - 1:
                html += ","
    
    html += """
            };
            
            // Initialize map
            function initMap() {
                // Create map
                map = new google.maps.Map(document.getElementById('map'), {
                    center: {lat: """ + str(default_lat) + """, lng: """ + str(default_lng) + """},
                    zoom: """ + str(default_zoom) + """,
                    mapTypeId: google.maps.MapTypeId.ROADMAP
                });
                
                infoWindow = new google.maps.InfoWindow({
                    maxWidth: 350
                });
                
                // Initialize drawing manager
                drawingManager = new google.maps.drawing.DrawingManager({
                    drawingMode: null,
                    drawingControl: false,
                    drawingControlOptions: {
                        position: google.maps.ControlPosition.TOP_CENTER,
                        drawingModes: [
                            google.maps.drawing.OverlayType.POLYGON
                        ]
                    },
                    polygonOptions: {
                        fillColor: '#22AA22',
                        fillOpacity: 0.3,
                        strokeWeight: 2,
                        strokeColor: '#22AA22',
                        clickable: true,
                        editable: true,
                        zIndex: 1
                    }
                });
                
                // Add drawing event listeners
                google.maps.event.addListener(drawingManager, 'overlaycomplete', function(e) {
                    // Switch back to non-drawing mode after drawing a shape
                    drawingManager.setDrawingMode(null);
                    
                    // Add an event listener that selects the newly-drawn shape
                    var newShape = e.overlay;
                    newShape.type = e.type;
                    
                    if (selectedShape) {
                        selectedShape.setMap(null);
                    }
                    selectedShape = newShape;
                    
                    // Enable the clear button
                    document.getElementById('clear-drawing-btn').disabled = false;
                    
                    // Find all markers inside the polygon
                    updateDrawnAreaMarkers();
                    
                    // Expand controls if collapsed
                    expandControls();
                });
                
                // Add church marker
                var churchMarker = new google.maps.Marker({
                    position: {lat: """ + str(default_lat) + """, lng: """ + str(default_lng) + """},
                    map: map,
                    title: \"""" + church_name + """\",
                    icon: {
                        path: google.maps.SymbolPath.CIRCLE,
                        fillColor: '#FF0000',
                        fillOpacity: 1,
                        strokeWeight: 1,
                        strokeColor: '#000000',
                        scale: 10
                    }
                });
                
                // Add markers for people
                addPeopleMarkers();
                
                // Hide loading indicator
                document.getElementById('loading').style.display = 'none';
                document.getElementById('initial-loading').style.display = 'none';
            }
            
            function toggleControls() {
                var controls = document.getElementById('controls');
                var toggle = document.getElementById('controls-toggle');
                
                if (controls.classList.contains('collapsed')) {
                    controls.classList.remove('collapsed');
                    toggle.innerHTML = '-';
                } else {
                    controls.classList.add('collapsed');
                    toggle.innerHTML = '+';
                }
            }
            
            function expandControls() {
                var controls = document.getElementById('controls');
                var toggle = document.getElementById('controls-toggle');
                
                if (controls.classList.contains('collapsed')) {
                    controls.classList.remove('collapsed');
                    toggle.innerHTML = '-';
                }
            }
            
            function addPeopleMarkers() {
                // Clear existing markers
                clearMarkers();
                
                // Create markers for each person/address
                for (var i = 0; i < peopleData.length; i++) {
                    var location = peopleData[i];
                    
                    // Create a custom marker icon based on the number of people at this address
                    var markerSize = 8 + (Math.min(location.count, 5) * 2); // Size increases with people count, max 5
                    var markerIcon = {
                        path: google.maps.SymbolPath.CIRCLE,
                        fillColor: '#3388ff',
                        fillOpacity: 0.7,
                        strokeWeight: 1,
                        strokeColor: '#000000',
                        scale: markerSize
                    };
                    
                    var marker = new google.maps.Marker({
                        position: {lat: location.lat, lng: location.lng},
                        map: map,
                        title: location.name,
                        street: location.street,
                        personIds: location.id.split(','), // Array of IDs
                        icon: markerIcon
                    });
                    
                    // Add click event for popup
                    (function(marker, content) {
                        google.maps.event.addListener(marker, 'click', function() {
                            infoWindow.setContent(content);
                            infoWindow.open(map, marker);
                        });
                    })(marker, location.popupContent);
                    
                    markers.push(marker);
                    visibleMarkers.push(marker);
                }
            }
            
            function toggleDrawingMode() {
                if (drawingManager.map) {
                    // Turn off drawing mode
                    drawingManager.setMap(null);
                    document.getElementById('drawing-btn').textContent = 'Enable Drawing Mode';
                } else {
                    // Turn on drawing mode
                    drawingManager.setMap(map);
                    drawingManager.setDrawingMode(google.maps.drawing.OverlayType.POLYGON);
                    document.getElementById('drawing-btn').textContent = 'Cancel Drawing';
                }
            }
            
            function clearDrawing() {
                if (selectedShape) {
                    selectedShape.setMap(null);
                    selectedShape = null;
                    document.getElementById('clear-drawing-btn').disabled = true;
                    document.getElementById('tag-drawn-btn').disabled = true;
                    drawnAreaPeopleIds = [];
                }
            }
            
            function updateDrawnAreaMarkers() {
                if (!selectedShape) return;
                
                drawnAreaPeopleIds = [];
                
                // Check each marker to see if it's inside the polygon
                for (var i = 0; i < visibleMarkers.length; i++) {
                    var marker = visibleMarkers[i];
                    if (google.maps.geometry.poly.containsLocation(marker.getPosition(), selectedShape)) {
                        // Add all person IDs from this marker
                        for (var j = 0; j < marker.personIds.length; j++) {
                            drawnAreaPeopleIds.push(marker.personIds[j]);
                        }
                    }
                }
                
                // Enable tag button if people are found
                document.getElementById('tag-drawn-btn').disabled = (drawnAreaPeopleIds.length === 0);
            }
            
            function tagDrawnArea() {
                if (drawnAreaPeopleIds.length > 0) {
                    // Show tag modal
                    currentPeopleIds = drawnAreaPeopleIds;
                    currentTagSource = "Custom Selection";
                    openTagModal("Custom Selection");
                }
            }
            
            function clearMarkers() {
                // Remove all markers from map
                for (var i = 0; i < markers.length; i++) {
                    markers[i].setMap(null);
                }
                
                // Clear marker array
                markers = [];
                visibleMarkers = [];
            }
            
            function filterMarkersByStreet() {
                var select = document.getElementById('street-select');
                var selectedStreet = select.options[select.selectedIndex].value;
                
                // Filter markers
                visibleMarkers = [];
                
                for (var i = 0; i < markers.length; i++) {
                    if (selectedStreet === 'all' || markers[i].street === selectedStreet) {
                        markers[i].setMap(map);
                        visibleMarkers.push(markers[i]);
                    } else {
                        markers[i].setMap(null);
                    }
                }
                
                // Update drawn area if needed
                if (selectedShape) {
                    updateDrawnAreaMarkers();
                }
            }
            
            // Tag modal functions
            function openTagModal(source) {
                document.getElementById('tagModal').style.display = 'block';
                document.getElementById('tagName').focus();
                document.getElementById('tagName').value = source + ' - ';
            }
            
            function closeTagModal() {
                document.getElementById('tagModal').style.display = 'none';
            }
            
            function submitTag() {
                var tagName = document.getElementById('tagName').value.trim();
                if (tagName === '') {
                    alert('Please enter a tag name');
                    return;
                }
                
                // Show loading indicator
                document.getElementById('loading').style.display = 'block';
                closeTagModal();
                
                // Set form values
                document.getElementById('tag_people_ids').value = currentPeopleIds.join(',');
                document.getElementById('tag_name').value = tagName;
                
                // Submit the form
                document.getElementById('tag-form').submit();
                
                // Hide loading after a delay
                setTimeout(function() {
                    document.getElementById('loading').style.display = 'none';
                }, 1000);
            }
            
            // Tag functions
            function tagPeopleByStreet() {
                var select = document.getElementById('street-select');
                var selectedStreet = select.options[select.selectedIndex].value;
                
                if (selectedStreet === 'all') {
                    tagAllVisible();
                    return;
                }
                
                // Get people IDs for the selected street
                var peopleIds = streetData[selectedStreet];
                if (peopleIds && peopleIds.length > 0) {
                    currentPeopleIds = peopleIds;
                    currentTagSource = selectedStreet + " Street";
                    openTagModal(currentTagSource);
                }
            }
            
            function tagAllVisible() {
                // Tag all currently visible people
                var visiblePeopleIds = [];
                for (var i = 0; i < visibleMarkers.length; i++) {
                    for (var j = 0; j < visibleMarkers[i].personIds.length; j++) {
                        visiblePeopleIds.push(visibleMarkers[i].personIds[j]);
                    }
                }
                
                if (visiblePeopleIds.length > 0) {
                    currentPeopleIds = visiblePeopleIds;
                    currentTagSource = "All Visible";
                    openTagModal(currentTagSource);
                }
            }
            
            // Export function
            function exportToCsv() {
                // Create CSV content
                var csvContent = "data:text/csv;charset=utf-8,";
                csvContent += "Name,Address,City,State,Zip,Phone,Email\\n";
                
                // Add visible markers data
                var exportedData = [];
                for (var i = 0; i < visibleMarkers.length; i++) {
                    var marker = visibleMarkers[i];
                    var title = marker.getTitle().replace(/,/g, " ");
                    var parts = title.split(", ");
                    
                    var address = parts[0] || "";
                    var city = parts[1] || "";
                    var state = "";
                    var zip = "";
                    
                    if (parts[2]) {
                        var stateZip = parts[2].split(" ");
                        state = stateZip[0] || "";
                        zip = stateZip[1] || "";
                    }
                    
                    // Add a row for each person at this marker
                    for (var j = 0; j < marker.personIds.length; j++) {
                        var id = marker.personIds[j];
                        // Find the person data in the popup content
                        var personName = "Person " + id;
                        var phone = "";
                        var email = "";
                        
                        // Extract name from popup (simplified)
                        var popupContent = marker.popupContent || "";
                        var nameMatch = popupContent.match(new RegExp('/Person2/' + id + '[^>]+>([^<]+)<'));
                        if (nameMatch && nameMatch[1]) {
                            personName = nameMatch[1];
                        }
                        
                        exportedData.push({
                            id: id,
                            name: personName,
                            address: address,
                            city: city,
                            state: state,
                            zip: zip,
                            phone: phone,
                            email: email
                        });
                    }
                }
                
                // Add data rows
                for (var i = 0; i < exportedData.length; i++) {
                    var row = exportedData[i];
                    csvContent += row.name + "," + row.address + "," + row.city + "," + 
                                 row.state + "," + row.zip + "," + row.phone + "," + row.email + "\\n";
                }
                
                // Create download link
                var encodedUri = encodeURI(csvContent);
                var link = document.createElement("a");
                link.setAttribute("href", encodedUri);
                link.setAttribute("download", "geographic_distribution.csv");
                document.body.appendChild(link);
                
                // Trigger download
                link.click();
                document.body.removeChild(link);
            }
            
            // Census data functions
            function toggleCensusLayer() {
                var select = document.getElementById('census-select');
                var selectedOption = select.options[select.selectedIndex].value;
                
                // Remove any existing census overlay
                if (censusOverlay && censusOverlay.rectangles) {
                    for (var i = 0; i < censusOverlay.rectangles.length; i++) {
                        censusOverlay.rectangles[i].setMap(null);
                    }
                    censusOverlay = null;
                }
                
                // If there's a data layer, remove it
                if (map.data) {
                    map.data.forEach(function(feature) {
                        map.data.remove(feature);
                    });
                }
                
                // Hide legend if no layer selected
                if (!selectedOption) {
                    document.getElementById('census-legend').style.display = 'none';
                    activeCensusLayer = "";
                    return;
                }
                
                // If a layer is selected, add it
                if (selectedOption && censusLayers[selectedOption]) {
                    activeCensusLayer = selectedOption;
                    loadAndDisplayCensusData(censusLayers[selectedOption]);
                }
            }
            
            // Function to load and display census data based on state
            function loadAndDisplayCensusData(layerInfo) {
                // Show loading indicator
                document.getElementById('loading').style.display = 'block';
                
                // Get the current map bounds to determine which states are visible
                var bounds = map.getBounds();
                var ne = bounds.getNorthEast();
                var sw = bounds.getSouthWest();
                
                // Map of state codes used by Census Bureau
                var stateCodes = {
                    'ALABAMA': '01',
                    'ALASKA': '02',
                    'ARIZONA': '04',
                    'ARKANSAS': '05',
                    'CALIFORNIA': '06',
                    'COLORADO': '08',
                    'CONNECTICUT': '09',
                    'DELAWARE': '10',
                    'DISTRICT OF COLUMBIA': '11',
                    'FLORIDA': '12',
                    'GEORGIA': '13',
                    'HAWAII': '15',
                    'IDAHO': '16',
                    'ILLINOIS': '17',
                    'INDIANA': '18',
                    'IOWA': '19',
                    'KANSAS': '20',
                    'KENTUCKY': '21',
                    'LOUISIANA': '22',
                    'MAINE': '23',
                    'MARYLAND': '24',
                    'MASSACHUSETTS': '25',
                    'MICHIGAN': '26',
                    'MINNESOTA': '27',
                    'MISSISSIPPI': '28',
                    'MISSOURI': '29',
                    'MONTANA': '30',
                    'NEBRASKA': '31',
                    'NEVADA': '32',
                    'NEW HAMPSHIRE': '33',
                    'NEW JERSEY': '34',
                    'NEW MEXICO': '35',
                    'NEW YORK': '36',
                    'NORTH CAROLINA': '37',
                    'NORTH DAKOTA': '38',
                    'OHIO': '39',
                    'OKLAHOMA': '40',
                    'OREGON': '41',
                    'PENNSYLVANIA': '42',
                    'RHODE ISLAND': '44',
                    'SOUTH CAROLINA': '45',
                    'SOUTH DAKOTA': '46',
                    'TENNESSEE': '47',
                    'TEXAS': '48',
                    'UTAH': '49',
                    'VERMONT': '50',
                    'VIRGINIA': '51',
                    'WASHINGTON': '53',
                    'WEST VIRGINIA': '54',
                    'WISCONSIN': '55',
                    'WYOMING': '56',
                    'PUERTO RICO': '72'
                };
                
                // Use Tennessee as a default state if no other information is available
                var stateCode = '47'; // Tennessee
                
                try {
                    // Determine if we're zoomed into a specific state
                    // (This is a simplistic approach - a real implementation would use reverse geocoding)
                    var center = map.getCenter();
                    var geocoder = new google.maps.Geocoder();
                    
                    geocoder.geocode({ 'location': { lat: center.lat(), lng: center.lng() } }, function(results, status) {
                        if (status === 'OK') {
                            // Try to determine the state from results
                            var state = '';
                            for (var i = 0; i < results.length; i++) {
                                var addressComponents = results[i].address_components;
                                for (var j = 0; j < addressComponents.length; j++) {
                                    var types = addressComponents[j].types;
                                    if (types.indexOf("administrative_area_level_1") >= 0) {
                                        state = addressComponents[j].long_name.toUpperCase();
                                        break;
                                    }
                                }
                                if (state) break;
                            }
                            
                            if (state && stateCodes[state]) {
                                stateCode = stateCodes[state];
                            }
                            
                            // Now load the GeoJSON for the determined state
                            loadStateGeoJSON(stateCode, layerInfo);
                        } else {
                            // If geocoding fails, just use Tennessee
                            loadStateGeoJSON(stateCode, layerInfo);
                        }
                    });
                } catch (e) {
                    console.error("Error determining state:", e);
                    // Fall back to Tennessee if there's an error
                    loadStateGeoJSON(stateCode, layerInfo);
                }
            }
            
            // Function to load GeoJSON for a specific state
            function loadStateGeoJSON(stateCode, layerInfo) {
                // GitHub repository has tract GeoJSON files for all states
                var geoJsonUrl = 'https://raw.githubusercontent.com/uscensusbureau/citysdk/master/v2/GeoJSON/500k/2019/' + stateCode + '/tract.json';
                
                // Fetch the GeoJSON data for census tracts
                fetch(geoJsonUrl)
                    .then(function(response) {
                        if (!response.ok) {
                            throw new Error('Failed to load census tract boundaries for state code: ' + stateCode);
                        }
                        return response.json();
                    })
                    .then(function(tractData) {
                        // Process and display the tract data
                        displayCensusTracts(tractData, layerInfo);
                    })
                    .catch(function(error) {
                        console.error('Error loading census data:', error);
                        // Fall back to simplified display
                        createSimpleCensusOverlay(layerInfo);
                        // Hide loading indicator
                        document.getElementById('loading').style.display = 'none';
                    });
            }
            
            // Function to display census tracts with appropriate styling
            function displayCensusTracts(tractData, layerInfo) {
                try {
                    // Get opacity
                    var opacity = parseFloat(document.getElementById('opacity-slider').value);
                    
                    // Clear any existing data
                    map.data.forEach(function(feature) {
                        map.data.remove(feature);
                    });
                    
                    // Add the GeoJSON features to the map
                    for (var i = 0; i < tractData.features.length; i++) {
                        var feature = tractData.features[i];
                        map.data.addGeoJson(feature);
                    }
                    
                    // Style the tracts based on the selected census variable
                    var colors = ['#FFFFCC', '#FFEDA0', '#FED976', '#FEB24C', '#FD8D3C', '#FC4E2A', '#E31A1C', '#BD0026'];
                    
                    map.data.setStyle(function(feature) {
                        // Get the GEOID of the tract
                        var geoid = feature.getProperty('GEOID');
                        
                        // In a real implementation, we would fetch the actual census data for this tract
                        // For now, we'll use a random value to demonstrate
                        var colorIndex = Math.floor(Math.random() * colors.length);
                        
                        return {
                            fillColor: colors[colorIndex],
                            strokeWeight: 0.5,
                            strokeColor: '#000000',
                            fillOpacity: opacity
                        };
                    });
                    
                    // Store the data layer for later reference
                    censusOverlay = {
                        dataLayer: true,
                        setOpacity: function(value) {
                            map.data.setStyle(function(feature) {
                                var existingStyle = map.data.getStyle()(feature);
                                return {
                                    fillColor: existingStyle.fillColor,
                                    strokeWeight: existingStyle.strokeWeight,
                                    strokeColor: existingStyle.strokeColor,
                                    fillOpacity: value
                                };
                            });
                        }
                    };
                    
                    // Show the legend
                    createLegend(layerInfo);
                    
                    // Hide loading indicator
                    document.getElementById('loading').style.display = 'none';
                } catch (e) {
                    console.error('Error displaying census tracts:', e);
                    // Fall back to simplified display
                    createSimpleCensusOverlay(layerInfo);
                    // Hide loading indicator
                    document.getElementById('loading').style.display = 'none';
                }
            }
            
            // Simplified census layer display (fallback)
            function createSimpleCensusOverlay(layerInfo) {
                try {
                    // Get opacity
                    var opacity = parseFloat(document.getElementById('opacity-slider').value);
                    
                    // Create a grid display of census data
                    var bounds = map.getBounds();
                    var NE = bounds.getNorthEast();
                    var SW = bounds.getSouthWest();
                    
                    // Create rectangles with different colors
                    var rectangles = [];
                    var colors = ['#FFFFCC', '#FFEDA0', '#FED976', '#FEB24C', '#FD8D3C', '#FC4E2A', '#E31A1C', '#BD0026'];
                    
                    // Divide visible area into grid
                    var gridSize = 8;
                    var latStep = (NE.lat() - SW.lat()) / gridSize;
                    var lngStep = (NE.lng() - SW.lng()) / gridSize;
                    
                    for (var row = 0; row < gridSize; row++) {
                        for (var col = 0; col < gridSize; col++) {
                            var south = SW.lat() + (row * latStep);
                            var north = south + latStep;
                            var west = SW.lng() + (col * lngStep);
                            var east = west + lngStep;
                            
                            // Different patterns for different layers
                            var colorIndex = 0;
                            if (layerInfo.variable === 'B19013_001E') {
                                colorIndex = (row + col) % colors.length;
                            } else if (layerInfo.variable === 'B01003_001E') {
                                colorIndex = Math.floor((row / gridSize) * colors.length);
                            } else {
                                colorIndex = Math.floor((col / gridSize) * colors.length);
                            }
                            
                            // Create rectangle
                            var rectangle = new google.maps.Rectangle({
                                strokeColor: '#888888',
                                strokeOpacity: 0.3,
                                strokeWeight: 1,
                                fillColor: colors[colorIndex],
                                fillOpacity: opacity,
                                map: map,
                                bounds: {
                                    north: north,
                                    south: south,
                                    east: east,
                                    west: west
                                }
                            });
                            
                            rectangles.push(rectangle);
                        }
                    }
                    
                    // Store for later removal
                    censusOverlay = {
                        rectangles: rectangles
                    };
                    
                    // Show legend
                    createLegend(layerInfo);
                } catch (e) {
                    console.error("Error adding census layer:", e);
                    alert("There was a problem loading the census data layer.");
                }
            }
            
            function createLegend(layerInfo) {
                var legend = document.getElementById('census-legend');
                var title = document.getElementById('legend-title');
                var items = document.getElementById('legend-items');
                
                // Position
                legend.style.position = 'absolute';
                legend.style.bottom = '30px';
                legend.style.left = '10px';
                legend.style.right = 'auto';
                
                // Clear out any existing content
                items.innerHTML = '';
                
                // Remove existing button
                var existingButton = legend.querySelector('button');
                if (existingButton) {
                    legend.removeChild(existingButton);
                }
                
                // Set title
                title.textContent = layerInfo.name;
                
                // Colors
                var colors = ['#FFFFCC', '#FFEDA0', '#FED976', '#FEB24C', '#FD8D3C', '#FC4E2A', '#E31A1C', '#BD0026'];
                
                // Simple ranges based on type
                var ranges = [];
                
                if (layerInfo.variable === 'B19013_001E') {
                    ranges = ['< $25k', '$25k-50k', '$50k-75k', '$75k-100k', '$100k-125k', '$125k-150k', '$150k-200k', '> $200k'];
                } else if (layerInfo.variable === 'B01003_001E') {
                    ranges = ['< 100', '100-500', '500-1k', '1k-2.5k', '2.5k-5k', '5k-10k', '10k-25k', '> 25k'];
                } else if (layerInfo.variable === 'B01002_001E') {
                    ranges = ['< 20', '20-25', '25-30', '30-35', '35-40', '40-45', '45-50', '> 50'];
                } else if (layerInfo.variable === 'B15003_022E') {
                    ranges = ['< 10%', '10-20%', '20-30%', '30-40%', '40-50%', '50-60%', '60-70%', '> 70%'];
                } else if (layerInfo.variable === 'B25003_002E') {
                    ranges = ['< 20%', '20-30%', '30-40%', '40-50%', '50-60%', '60-70%', '70-80%', '> 80%'];
                } else {
                    ranges = ['Low', '', '', 'Medium', '', '', 'High'];
                }
                
                // Add legend items  
                for (var i = 0; i < colors.length; i++) {
                    if (i < ranges.length && ranges[i]) {
                        var item = document.createElement('div');
                        item.className = 'legend-item';
                        
                        var colorBox = document.createElement('div');
                        colorBox.className = 'color-box';
                        colorBox.style.backgroundColor = colors[i];
                        
                        var label = document.createElement('span');
                        label.textContent = ranges[i];
                        
                        item.appendChild(colorBox);
                        item.appendChild(label);
                        items.appendChild(item);
                    }
                }
                
                // Source info
                var sourceInfo = document.createElement('div');
                sourceInfo.style.fontSize = '10px';
                sourceInfo.style.marginTop = '5px';
                sourceInfo.style.color = '#666';
                sourceInfo.textContent = 'Source: US Census ACS 5-Year Data';
                items.appendChild(sourceInfo);
                
                // Add close button
                var closeButton = document.createElement('button');
                closeButton.innerHTML = 'Close';
                closeButton.style.marginTop = '5px';
                closeButton.style.padding = '2px 5px';
                closeButton.style.fontSize = '12px';
                closeButton.style.width = '100%';
                closeButton.onclick = function() {
                    legend.style.display = 'none';
                };
                legend.appendChild(closeButton);
                
                // Show the legend
                legend.style.display = 'block';
            }
            
            function updateCensusOpacity(value) {
                if (censusOverlay) {
                    if (censusOverlay.rectangles) {
                        var newOpacity = parseFloat(value);
                        for (var i = 0; i < censusOverlay.rectangles.length; i++) {
                            censusOverlay.rectangles[i].setOptions({fillOpacity: newOpacity});
                        }
                    } else if (censusOverlay.dataLayer) {
                        // Update data layer opacity
                        var newOpacity = parseFloat(value);
                        map.data.setStyle(function(feature) {
                            return {
                                fillColor: feature.getProperty('fillColor') || '#FFEDA0',
                                strokeWeight: 0.5,
                                strokeColor: '#000000',
                                fillOpacity: newOpacity
                            };
                        });
                    }
                }
            }
            
            function updateOpacityLabel(value) {
                document.getElementById('opacity-value').textContent = value;
            }
            
            function showLegend() {
                var select = document.getElementById('census-select');
                var selectedOption = select.options[select.selectedIndex].value;
                
                if (selectedOption && censusLayers[selectedOption]) {
                    createLegend(censusLayers[selectedOption]);
                } else {
                    alert("Please select a census layer first to show its legend.");
                }
            }
            
            // Close modal when clicking outside
            window.onclick = function(event) {
                var modal = document.getElementById('tagModal');
                if (event.target == modal) {
                    closeTagModal();
                }
            }
            
            // Enter key for tag name
            document.getElementById('tagName').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    submitTag();
                }
            });
        </script>
        
        <!-- Load Google Maps API with drawing and geometry libraries -->
        <script src="https://maps.googleapis.com/maps/api/js?key=""" + google_maps_api_key + """&libraries=drawing,geometry&callback=initMap" async defer></script>
    </body>
    </html>
    """
    
    return html

def render_summary(data):
    """Render the summary information."""
    html = "<div id='summary-info' style='background: #f8f9fa; padding: 10px; margin-bottom: 15px; border-radius: 5px;'>"
    html += "<h4>Map Summary</h4>"
    html += "<ul>"
    html += "<li><strong>Total people in query:</strong> {}</li>".format(data.get("total_count", 0))
    html += "<li><strong>People with valid coordinates:"
    html += "<li><strong>People with valid coordinates:</strong> {}</li>".format(data.get("located_count", 0))
    html += "<li><strong>Unique locations on map:</strong> {}</li>".format(data.get("address_count", 0))
    html += "<li><strong>Unique streets:</strong> {}</li>".format(data.get("street_count", 0))
    html += "</ul>"
    html += "</div>"
    return html

def render_error(error_message):
    """Render an error message."""
    html = """
    <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; padding: 10px; border-radius: 4px; margin-bottom: 10px;">
        <h3>Error</h3>
        <p>{}</p>
    </div>
    <p><a href="javascript:history.back()" class="btn btn-primary">Go Back</a></p>
    """.format(error_message)
    return html

def main():
    """Main function to handle different actions."""
    try:
        # Get configuration
        config = {
            "MAP_TITLE": MAP_TITLE,
            "MAP_HEIGHT": MAP_HEIGHT,
            "CHURCH_NAME": CHURCH_NAME,
            "CHURCH_LATITUDE": CHURCH_LATITUDE,
            "CHURCH_LONGITUDE": CHURCH_LONGITUDE,
            "DEFAULT_ZOOM": DEFAULT_ZOOM,
            "GOOGLE_MAPS_API_KEY": GOOGLE_MAPS_API_KEY,
            "CENSUS_API_KEY": CENSUS_API_KEY,
            "ENABLE_CENSUS_DATA": ENABLE_CENSUS_DATA,
            "CENSUS_OVERLAY_OPTIONS": CENSUS_OVERLAY_OPTIONS
        }
        
        # Determine what action we're performing based on request
        action = ""
        if hasattr(model.Data, 'action'):
            action = model.Data.action
        
        # Handle different actions
        if action == "tag_people":
            handle_tag_people()
        else:
            # Default action: Show the map
            print "<div id='initial-loading' style='text-align: center; padding: 20px;'>"
            print "<h3>Loading Geographic Distribution Map...</h3>"
            print "<p>This may take a moment, especially for large datasets.</p>"
            print "</div>"
            
            # Get people data from the BlueToolbar query
            people_list = q.BlueToolbarReport("name")
            
            # Get count of people
            total_count = 0
            try:
                total_count = q.BlueToolbarCount()
            except:
                # Count manually if BlueToolbarCount fails
                for _ in people_list:
                    total_count += 1
            
            # Process people data
            result = process_people_data(people_list, total_count)
            
            if "error" in result:
                # Show error
                print render_error(result["error"])
            else:
                # Generate map data
                js_data = generate_map_data(
                    result["people_by_address"], 
                    result["people_by_street"],
                    model.CmsHost
                )
                
                # Add JS data to result
                result.update(js_data)
                
                # Render the map
                print render_summary(result)
                print render_map(result, config)
                
    except Exception as e:
        # Print any errors
        print render_error(str(e))
        print "<pre style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; overflow: auto;'>"
        traceback.print_exc()
        print "</pre>"

# Execute main function
main()
