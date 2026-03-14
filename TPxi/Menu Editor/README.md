# Menu Editor

A visual UI for managing TouchPoint's XML menu configurations stored in Special Content (Text Content). Easily add, edit, reorder, and delete menu items across all report menus and the Blue Toolbar — without touching raw XML.

With this tool, admins can:

- Visually manage all 5 XML menu configuration files from one interface
- Add, edit, delete, and reorder headers and report items with drag-and-drop
- Move items between columns with arrow controls
- Set role-based permissions using a dropdown populated from your TouchPoint roles
- Manage Blue Toolbar entries grouped by type (SQL Reports, Python Scripts, Other Reports, Custom Reports)
- Edit column definitions for custom data export reports
- Detect duplicate report names in the Blue Toolbar with highlighted warnings
- Switch to raw XML for advanced manual editing
- Discard unsaved changes across any or all tabs
- Automatic backup before every save with full audit trail
- Preview and one-click restore from backup history

---

## Reports Menu Editor

Manage the 4-column layout for Admin, Finance, Involvements, and People report menus. Headers and reports are displayed as cards with drag-and-drop reordering, column management, and role assignment.

<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Menu%20Editor/ME-1.png" width="700" alt="Reports Menu Editor - People Tab">
</p>

---

## Blue Toolbar Editor

Manage CustomReports entries grouped by type — SQL Reports, Python Scripts, Other Reports, and Custom Reports with column definitions. Duplicate report names are automatically detected and highlighted with a warning banner.

<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Menu%20Editor/ME-2.png" width="700" alt="Blue Toolbar Editor with Duplicate Detection">
</p>

---

## Managed XML Files

| Tab | Content Name | Description |
|-----|-------------|-------------|
| Admin | ReportsMenuAdmin | Admin reports menu |
| Finance | ReportsMenuFinance | Finance reports menu |
| Involvements | ReportsMenuInvolvements | Involvements reports menu |
| People | ReportsMenuPeople | People reports menu |
| Blue Toolbar | CustomReports | Blue toolbar and custom report definitions |

---

## Setup

1. Navigate to **Admin > Advanced > Special Content > Python**
2. Click **New Python Script File**
3. Name it `TPxi_MenuEditor`
4. Paste the contents of `TPxi_MenuEditor.py`
5. Add to CustomReports XML:
   ```xml
   <Report name="TPxi_MenuEditor" type="PyScript" role="Admin" />
   ```

---

## Backup System

Backups are created automatically before every save. Up to 10 backups are kept per content file, each logging who made the change and when. Backups can be previewed and restored with one click, and a safety backup of the current content is created before any restore.

---

*Like this tool? [DisplayCache](https://displaycache.com) integrates directly with TouchPoint and helps fund continued development of tools like this one.*
