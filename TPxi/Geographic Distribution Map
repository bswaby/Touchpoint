"""
Geographic Distribution Map with Tagging Feature
---------------------------
This script creates an interactive map showing the geographic distribution of people
from a BlueToolbar query, with profile photos, grouped markers for same address,
and ability to tag people from selections.

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

# Import necessary libraries
import traceback
import re
import collections

try:
    # Determine what action we're performing based on request
    action = ""
    if hasattr(model.Data, 'action'):
        action = model.Data.action

    # Configuration settings
    MAP_TITLE = "Geographic Distribution Map"
    DEFAULT_LATITUDE = 36.3204292  # Church address: 106 Bluegrass Commons Blvd, Hendersonville, TN
    DEFAULT_LONGITUDE = -86.5681495  # Church address: 106 Bluegrass Commons Blvd, Hendersonville, TN
    DEFAULT_ZOOM = 11  # Default zoom level
    MAP_HEIGHT = 800  # Map height in pixels
    GOOGLE_MAPS_API_KEY = ""  # Default key
    
    # Helper function to safely convert string to float
    def safe_float(value):
        try:
            if value is None:
                return None
            return float(value)
        except:
            return None
    
    # Main execution based on action parameter
    if action == "tag_people":
        # Check if form data is available
        if hasattr(model.Data, 'people_ids') and hasattr(model.Data, 'tag_name'):
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
        else:
            # No data provided
            print "<h2>Tagging Error</h2>"
            print "<p>No people selected for tagging. Please return to the map and select people to tag.</p>"
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
        people_by_street = {}
        located_count = 0
        
        # For grouping by address
        people_by_address = collections.defaultdict(list)
        
        # Get people IDs from the BlueToolbar query
        people_ids = []
        for person in people_list:
            people_ids.append(str(person.PeopleId))
        
        # Get geocode data using direct query (hidden processing)
        if len(people_ids) > 0:
            try:
                # Create a simple Person class to hold the data
                class SimplePerson:
                    def __init__(self, pid, name, address, city, state, zip_code, home_phone, cell_phone, email, lat, lng, picture_url=None):
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
                        # Create a full address for grouping
                        self.FullAddress = (address or "") + ", " + (city or "") + ", " + (state or "")
                
                # Direct query based on your screenshot, with picture info - corrected join condition
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
                
                # Debug info - hidden by default but can be uncommented for troubleshooting
                # print "<div style='display:none;'><p>Debug SQL: {}</p></div>".format(direct_sql.replace("<", "&lt;").replace(">", "&gt;"))
                
                for row in q.QuerySql(direct_sql):
                    try:
                        # Get latitude and longitude
                        lat = safe_float(row.Latitude)
                        lng = safe_float(row.Longitude)
                        
                        # Skip people without valid coordinates
                        if lat is None or lng is None:
                            continue
                        
                        # Get picture URL if available, clean up URL if needed
                        picture_url = getattr(row, 'PictureUrl', None)
                        if picture_url:
                            # Remove any domain prefixes
                            if picture_url.startswith('http'):
                                if '//img' in picture_url:
                                    picture_url = 'https://img' + picture_url.split('//img')[1]
                                elif 'https://' in picture_url and '/' in picture_url[8:]:
                                    parts = picture_url.split('https://')
                                    if len(parts) > 1:
                                        picture_url = 'https://' + parts[1]
                        
                        # Create a simple person object with the data we need
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
                        street = ""
                        if person.PrimaryAddress:
                            street_parts = person.PrimaryAddress.split(',')
                            if not street_parts:
                                street_parts = [person.PrimaryAddress]
                            street = re.sub(r'^\d+\s+', '', street_parts[0]).strip()
                        
                        # Add to street collection
                        if street:
                            if street not in people_by_street:
                                people_by_street[street] = []
                            people_by_street[street].append(person)
                        
                        # Group by full address for popup grouping
                        addr_key = (person.PrimaryAddress, person.PrimaryCity, person.PrimaryState)  # Use address as key for proper grouping
                        people_by_address[addr_key].append(person)
                        
                        located_count += 1
                        
                    except Exception as e:
                        continue
                
            except Exception as e:
                print "<div style='color: red;'>Error with query: {}</div>".format(str(e))
        
        # Show summary info
        print "<div id='summary-info' style='background: #f8f9fa; padding: 10px; margin-bottom: 15px; border-radius: 5px;'>"
        print "<h4>Map Summary</h4>"
        print "<ul>"
        print "<li><strong>Total people in query:</strong> {}</li>".format(total_count)
        print "<li><strong>People with valid coordinates:</strong> {}</li>".format(located_count)
        print "<li><strong>Unique locations on map:</strong> {}</li>".format(len(people_by_address))
        print "<li><strong>Unique streets:</strong> {}</li>".format(len(people_by_street))
        print "</ul>"
        print "</div>"
        
        # Create the JavaScript data for grouped markers
        people_data_js = "["
        marker_count = 0
        
        for addr_key, people_list in people_by_address.items():
            try:
                # All people at this address have the same coordinates - use the first person
                first_person = people_list[0]
                lat = safe_float(first_person.Latitude)
                lng = safe_float(first_person.Longitude)
                
                if lat is None or lng is None:
                    continue
                
                # Get street from first person
                street = ""
                if first_person.PrimaryAddress:
                    street_parts = first_person.PrimaryAddress.split(',')
                    if not street_parts:
                        street_parts = [first_person.PrimaryAddress]
                    street = re.sub(r'^\d+\s+', '', street_parts[0]).strip()
                
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
                for person in people_list:
                    popup_html += '<div class="person-entry">'
                    
                    # Add picture if available
                    if person.PictureUrl:
                        popup_html += '<img src="' + person.PictureUrl + '" class="popup-photo" />'
                    
                    popup_html += '<h4><a href="{}/Person2/{}" target="_blank">{}</a></h4>'.format(
                        model.CmsHost, person.PeopleId, person.Name)
                    
                    # Add contact info
                    if person.HomePhone:
                        popup_html += '<p><strong>Home:</strong> {}</p>'.format(
                            model.FmtPhone(person.HomePhone))
                    if person.EmailAddress:
                        popup_html += '<p><strong>Email:</strong> {}</p>'.format(person.EmailAddress)
                    
                    popup_html += '</div>'
                    if person != people_list[-1]:  # Add separator if not the last person
                        popup_html += '<hr class="person-separator" />'
                
                popup_html += '</div>'
                
                # Add person IDs for this address
                person_ids_str = ""
                for person in people_list:
                    person_ids_str += str(person.PeopleId) + ","
                person_ids_str = person_ids_str.rstrip(',')  # Remove trailing comma
                
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
        
        # Create street data for JavaScript
        street_data_js = "{"
        for street in people_by_street:
            street_data_js += "'" + street.replace("'", "\\'") + "': ["
            for person in people_by_street[street]:
                street_data_js += str(person.PeopleId) + ","
            street_data_js += "],"
        street_data_js += "};"
        
        # Get script name from URL path for proper form action
        script_name = "GeographicDistributionMap"  # Default name
        if hasattr(model, 'ScriptName'):
            script_name = model.ScriptName
        
        # Output the HTML
        html = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                #map-container { position: relative; height: """ + str(MAP_HEIGHT) + """px; width: 100%; }
                #map { height: 100%; width: 100%; }
                #controls { position: absolute; top: 10px; right: 10px; z-index: 1000; 
                          background: white; padding: 10px; border-radius: 4px; box-shadow: 0 0 10px rgba(0,0,0,0.2);
                          transition: height 0.3s ease; overflow: hidden; }
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
                
                /* Modal styles */
                .modal { display: none; position: fixed; z-index: 2001; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.4); }
                .modal-content { background-color: #fefefe; margin: 15% auto; padding: 20px; border: 1px solid #888; width: 300px; border-radius: 5px; }
                .modal-header { margin-bottom: 15px; }
                .modal-footer { margin-top: 15px; text-align: right; }
                .modal-footer button { margin-left: 10px; }
            </style>
        </head>
        <body>
            <h2>""" + MAP_TITLE + """<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg></h2>
            <div id="status-bar">
                Showing <strong>""" + str(located_count) + " of " + str(total_count) + """</strong> people on map at """ + str(marker_count) + """ locations. 
                <span style="color:#777">(People without valid addresses cannot be displayed)</span>
            </div>
            <div id="map-container">
                <div id="map"></div>
                <div id="controls" class="collapsed">
                    <div id="controls-header" onclick="toggleControls()">
                        <h4>Tag</h4>
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
        
        # Add street options to selector (limit to top 100 streets)
        street_count = 0
        for street in sorted(people_by_street.keys()):
            if street and street_count < 100:  # Limit to 100 streets
                count = len(people_by_street[street])
                html += '<option value="' + street.replace('"', '\\"') + '">' + street + ' (' + str(count) + ')</option>'
                street_count += 1
        
        html += """
                            </select>
                            <button onclick="tagPeopleByStreet()">Tag</button>
                        </div>
                        <button class="export-btn" onclick="tagAllVisible()">Tag All Visible</button>
                    </div>
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
                var peopleData = """ + people_data_js + """
                var streetData = """ + street_data_js + """
                var markers = [];
                var visibleMarkers = [];
                var map;
                var infoWindow;
                var drawingManager;
                var selectedShape;
                var drawnAreaPeopleIds = [];
                var currentPeopleIds = []; // For current selection to tag
                var currentTagSource = ""; // Source of the current selection for tag modal title
                
                // Initialize map
                function initMap() {
                    // Create map
                    map = new google.maps.Map(document.getElementById('map'), {
                        center: {lat: """ + str(DEFAULT_LATITUDE) + """, lng: """ + str(DEFAULT_LONGITUDE) + """},
                        zoom: """ + str(DEFAULT_ZOOM) + """,
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
                        position: {lat: """ + str(DEFAULT_LATITUDE) + """, lng: """ + str(DEFAULT_LONGITUDE) + """},
                        map: map,
                        title: "First Baptist Church",
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
                
                // Close modal when clicking outside
                window.onclick = function(event) {
                    var modal = document.getElementById('tagModal');
                    if (event.target == modal) {
                        closeTagModal();
                    }
                }
                
                // Add keypress handler for tag name input
                document.getElementById('tagName').addEventListener('keypress', function(e) {
                    if (e.key === 'Enter') {
                        submitTag();
                    }
                });
            </script>
            
            <!-- Load Google Maps API with drawing and geometry libraries -->
            <script src="https://maps.googleapis.com/maps/api/js?key=""" + GOOGLE_MAPS_API_KEY + """&libraries=drawing,geometry&callback=initMap" async defer></script>
        </body>
        </html>
        """
        
        # Output the map HTML
        print html

except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<div style='background-color: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; padding: 10px; border-radius: 4px; margin-bottom: 10px;'>"
    print "<p><strong>An error occurred:</strong> " + str(e) + "</p>"
    print "</div>"
    print "<p>Details:</p>"
    print "<pre style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; overflow: auto;'>"
    traceback.print_exc()
    print "</pre>"
    print "<p><a href='javascript:history.back()' class='btn btn-primary'>Go Back</a></p>"
