#----------------------------------------------------------------------
# TPxi_DayOfRegistration.py
#
# Day-of Registration Tool for TouchPoint
#
# Handles day-of registration at large events (VBS, camps, etc.)
# where multiple iPad stations are staffed by volunteers simultaneously.
# Each station handles a specific group (e.g., Kindergarten, 1st Grade)
# and assigns people to destination involvements with capacity limits.
#
# Features:
#   - Admin mode: Create/edit scenarios with stations and destinations
#   - Station mode: iPad-friendly registration interface
#   - Real-time sync: 10-second polling keeps all stations in sync
#   - Capacity management: Soft/hard caps on destination involvements
#   - Walk-in support: Global people search for unregistered attendees
#   - Friend finder: Search which destination a friend was assigned to
#   - Undo support: Reverse assignments with audit logging
#
#----------------------------------------------------------------------
# TPxi_RollSheet.py
#
# Configurable Rollsheet Generator for TouchPoint
#
# Allows admins to create reusable rollsheet configurations through a UI,
# then generate and print rollsheets for any ministry area. Each config
# specifies source (program/division or specific orgs), columns to display,
# universal data sources (reg questions, extra values, recreg), layout, and sort options.
#
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint  (40+ free tools)
# ----------------------------------------------------------------
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to
# support continued development, check out:
#
# DisplayCache - church digital signage that integrates with TouchPoint
# https://displaycache.com
#
# TPxi Go - your church contacts, wherever you work.
# Look up anyone in TouchPoint, log calls and emails from Outlook
# or your phone. No tab switching, no lost context.
# https://tpxigo.com
# ----------------------------------------------------------------
#
# Architecture:
#   Single .py file with two modes (admin vs station) in one script.
#   Python AJAX handlers for POST requests, HTML SPA for GET.
#   Data stored via model.WriteContentText / model.TextContent.
#
# Storage Keys:
#   DayOfRegistration_Scenarios - All scenario configs (JSON)
#   DayOfRegistration_Log_<id>  - Per-scenario processing log (JSON)
#
# CSS Prefix: dr-
# Root Class: .dr-root
#
# Reference:
#   TPxi_InvolvementProcessor.py - AJAX dispatch, JoinOrg patterns
#   FastLaneCheckIn.py - iPad-optimized CSS patterns
#----------------------------------------------------------------------

import json
import datetime
import random
import re

# --- Version / Auto-update -------------------------------------------
APP_VERSION = '1.1.4'
DC_SCRIPT_ID = 'TPxi_DayOfRegistration'
# scripts.displaycache.com is the custom domain used for browser-side version checks.
# workers.dev is used for server-side fetches (bypasses Cloudflare Bot Fight Mode).
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'

def get_script_name():
    """Detect the actual script name TouchPoint installed this as (admin may have
    renamed it). Order: posted script_name, model.URL parse, hardcoded default."""
    try:
        if hasattr(Data, 'script_name') and Data.script_name:
            sn = str(Data.script_name).strip()
            if sn:
                return sn
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID

model.Header = 'Day of Registration'

# =====================================================================
# AJAX HANDLERS (POST)
# =====================================================================
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -----------------------------------------------------------------
    # Helper: read the per-scenario assignment log and return ONLY the
    # (peopleId, orgId) pairs that this script actually assigned and
    # have not since been undone. Excludes any pre-existing room rosters,
    # leaders, or members added outside this script.
    # -----------------------------------------------------------------
    def _script_assigned_pairs(scenario_id):
        log_key = "DayOfRegistration_Log_" + scenario_id
        try:
            log_raw = model.TextContent(log_key) or ''
        except:
            log_raw = ''
        if not log_raw:
            return []
        try:
            log_data = json.loads(log_raw)
        except:
            return []
        # Walk the log in order. assign-then-undo cancels out.
        state_map = {}
        for e in log_data.get('entries', []):
            try:
                pid = int(e.get('peopleId') or 0)
                oid = int(e.get('destOrgId') or 0)
            except:
                continue
            if not pid or not oid:
                continue
            act = e.get('dr_action', '')
            if act == 'assign':
                state_map[(pid, oid)] = {
                    'action': 'assign',
                    'name': e.get('personName', ''),
                    'orgName': e.get('destOrgName', ''),
                    'displayName': e.get('destDisplayName', ''),
                    'stationName': e.get('stationName', '')
                }
            elif act == 'undo':
                state_map.pop((pid, oid), None)
        result = []
        for (pid, oid), info in state_map.items():
            if info.get('action') == 'assign':
                result.append({
                    'peopleId': pid,
                    'orgId': oid,
                    'name': info.get('name', ''),
                    'orgName': info.get('orgName', ''),
                    'displayName': info.get('displayName', ''),
                    'stationName': info.get('stationName', '')
                })
        return result

    # -----------------------------------------------------------------
    # ADMIN: Load all scenarios
    # -----------------------------------------------------------------
    if action == 'load_scenarios':
        try:
            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}
            print json.dumps({'success': True, 'scenarios': data.get('scenarios', [])})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Save scenario
    # -----------------------------------------------------------------
    elif action == 'save_scenario':
        try:
            scenario_json = str(Data.scenario_json) if hasattr(Data, 'scenario_json') else ''
            scenario = json.loads(scenario_json)

            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}
            scenarios = data.get('scenarios', [])

            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

            existing_idx = -1
            for i, s in enumerate(scenarios):
                if s.get('id') == scenario.get('id'):
                    existing_idx = i
                    break

            if existing_idx >= 0:
                scenario['updatedAt'] = now
                scenario['createdAt'] = scenarios[existing_idx].get('createdAt', now)
                scenarios[existing_idx] = scenario
            else:
                if not scenario.get('id'):
                    scenario['id'] = 'scenario_' + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + str(random.randint(100, 999))
                scenario['createdAt'] = now
                scenario['updatedAt'] = now
                scenarios.append(scenario)

            data['scenarios'] = scenarios
            model.WriteContentText("DayOfRegistration_Scenarios", json.dumps(data), "")
            print json.dumps({'success': True, 'scenario': scenario, 'message': 'Scenario saved'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Delete scenario
    # -----------------------------------------------------------------
    elif action == 'delete_scenario':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}
            scenarios = data.get('scenarios', [])

            new_scenarios = [s for s in scenarios if s.get('id') != scenario_id]
            if len(new_scenarios) < len(scenarios):
                data['scenarios'] = new_scenarios
                model.WriteContentText("DayOfRegistration_Scenarios", json.dumps(data), "")
                print json.dumps({'success': True, 'message': 'Scenario deleted'})
            else:
                print json.dumps({'success': False, 'message': 'Scenario not found'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Search involvements for source/destination config
    # -----------------------------------------------------------------
    elif action == 'search_involvements':
        try:
            search_term = getattr(Data, 'search_term', '')
            if search_term:
                search_term = str(search_term)
            division_id = getattr(Data, 'division_id', '')
            program_id = getattr(Data, 'program_id', '')

            where_clauses = ["o.OrganizationStatusId = 30"]

            if search_term:
                safe_term = search_term.replace("'", "''")
                where_clauses.append("""(o.OrganizationName LIKE '%{0}%'
                    OR CAST(o.OrganizationId AS VARCHAR) = '{0}')""".format(safe_term))

            if division_id:
                where_clauses.append("o.DivisionId = {0}".format(int(division_id)))

            if program_id:
                where_clauses.append("d.ProgId = {0}".format(int(program_id)))

            sql = """
                SELECT TOP 50
                    o.OrganizationId,
                    o.OrganizationName,
                    d.Name as DivisionName,
                    p.Name as ProgramName,
                    (SELECT COUNT(*) FROM OrganizationMembers om
                     WHERE om.OrganizationId = o.OrganizationId) as MemberCount
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE {0}
                ORDER BY o.OrganizationName
            """.format(" AND ".join(where_clauses))

            results = q.QuerySql(sql)
            involvements = []
            for r in results:
                involvements.append({
                    'orgId': r.OrganizationId,
                    'orgName': r.OrganizationName,
                    'divisionName': r.DivisionName or '',
                    'programName': r.ProgramName or '',
                    'memberCount': r.MemberCount or 0
                })
            print json.dumps({'success': True, 'involvements': involvements})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Get filter dropdowns (programs/divisions)
    # -----------------------------------------------------------------
    elif action == 'get_filters':
        try:
            prog_sql = "SELECT Id, Name FROM Program ORDER BY Name"
            div_sql = "SELECT Id, Name, ProgId FROM Division ORDER BY Name"

            programs = []
            for r in q.QuerySql(prog_sql):
                programs.append({'id': r.Id, 'name': r.Name})

            divisions = []
            for r in q.QuerySql(div_sql):
                divisions.append({'id': r.Id, 'name': r.Name, 'progId': r.ProgId})

            print json.dumps({'success': True, 'programs': programs, 'divisions': divisions})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Get subgroups for a target org
    # -----------------------------------------------------------------
    elif action == 'get_org_subgroups':
        try:
            org_id = int(str(Data.org_id)) if hasattr(Data, 'org_id') else 0
            sql = """
                SELECT DISTINCT sg.Name
                FROM MemberTags sg
                WHERE sg.OrgId = {0}
                ORDER BY sg.Name
            """.format(org_id)
            results = q.QuerySql(sql)
            subgroups = []
            for r in results:
                subgroups.append(r.Name)
            print json.dumps({'success': True, 'subgroups': subgroups})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Load station data (source registrants + dest counts)
    # -----------------------------------------------------------------
    elif action == 'load_station_data':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            station_id = str(Data.station_id) if hasattr(Data, 'station_id') else ''

            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}

            scenario = None
            for s in data.get('scenarios', []):
                if s.get('id') == scenario_id:
                    scenario = s
                    break

            if not scenario:
                print json.dumps({'success': False, 'message': 'Scenario not found'})
            else:
                station = None
                for st in scenario.get('stations', []):
                    if st.get('id') == station_id:
                        station = st
                        break

                if not station:
                    print json.dumps({'success': False, 'message': 'Station not found'})
                else:
                    source_org_id = station.get('sourceOrgId', 0)
                    registrants = []

                    if source_org_id:
                        sql = """
                            SELECT p.PeopleId,
                                   ISNULL(p.Name2, '') as Name2,
                                   ISNULL(p.FirstName, '') as FirstName,
                                   ISNULL(p.LastName, '') as LastName,
                                   ISNULL(p.NickName, '') as NickName,
                                   p.Age,
                                   ISNULL(g.Description, '') as Gender
                            FROM OrganizationMembers om
                            JOIN People p ON om.PeopleId = p.PeopleId
                            LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                            WHERE om.OrganizationId = {0}
                            ORDER BY p.LastName, p.FirstName
                        """.format(int(source_org_id))
                        results = q.QuerySql(sql)
                        for r in results:
                            try:
                                age_val = r.Age
                            except:
                                age_val = None
                            registrants.append({
                                'peopleId': r.PeopleId,
                                'name': str(r.Name2),
                                'firstName': str(r.FirstName),
                                'lastName': str(r.LastName),
                                'nickName': str(r.NickName),
                                'age': str(age_val) if age_val else '',
                                'gender': str(r.Gender)
                            })

                    destinations = station.get('destinations', [])
                    dest_counts = []
                    for dest in destinations:
                        dest_org_id = dest.get('orgId', 0)
                        count = 0
                        sg_counts = {}
                        if dest_org_id:
                            cnt_sql = "SELECT Count(om.PeopleId) as cnt FROM OrganizationMembers om WHERE om.OrganizationId = {0}".format(int(dest_org_id))
                            cnt_row = q.QuerySqlTop1(cnt_sql)
                            count = cnt_row.cnt if cnt_row and hasattr(cnt_row, 'cnt') else 0
                            # Get subgroup member counts
                            dest_sgs = dest.get('subgroups', [])
                            if dest_sgs:
                                sg_sql = "SELECT sg.Name, COUNT(DISTINCT ommt.PeopleId) as cnt FROM MemberTags sg LEFT JOIN OrgMemMemTags ommt ON sg.Id = ommt.MemberTagId WHERE sg.OrgId = {0} GROUP BY sg.Id, sg.Name".format(int(dest_org_id))
                                for sgr in q.QuerySql(sg_sql):
                                    if sgr.Name in dest_sgs:
                                        sg_counts[sgr.Name] = sgr.cnt or 0
                        dest_counts.append({
                            'destId': dest.get('id', ''),
                            'orgId': dest_org_id,
                            'displayName': dest.get('displayName', ''),
                            'orgName': dest.get('orgName', ''),
                            'subgroup': dest.get('subgroup', ''),
                            'subgroups': dest.get('subgroups', []),
                            'capacity': dest.get('capacity', 0),
                            'capType': dest.get('capType', 'soft'),
                            'currentCount': count,
                            'subgroupCounts': sg_counts
                        })

                    # Check which registrants are already assigned (query actual DB membership)
                    assigned_ids = {}
                    dest_org_ids = [d.get('orgId', 0) for d in destinations if d.get('orgId', 0)]
                    if dest_org_ids:
                        # Build display name map from destination config
                        dest_display_map = {}
                        for d in destinations:
                            if d.get('orgId'):
                                dest_display_map[d.get('orgId')] = d.get('displayName', '') or d.get('orgName', '')
                        assign_sql = """
                            SELECT om.PeopleId, om.OrganizationId,
                                   o.OrganizationName,
                                   ISNULL(p.Name2, '') as Name2
                            FROM OrganizationMembers om
                            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                            JOIN People p ON om.PeopleId = p.PeopleId
                            WHERE om.OrganizationId IN ({0})
                        """.format(','.join(str(x) for x in dest_org_ids))
                        for row in q.QuerySql(assign_sql):
                            pid = row.PeopleId
                            oid = row.OrganizationId
                            assigned_ids[pid] = {
                                'destOrgId': oid,
                                'destOrgName': str(row.OrganizationName or ''),
                                'destDisplayName': dest_display_map.get(oid, str(row.OrganizationName or '')),
                                'personName': str(row.Name2 or '')
                            }

                    print json.dumps({
                        'success': True,
                        'registrants': registrants,
                        'destinations': dest_counts,
                        'assignedIds': assigned_ids,
                        'stationName': station.get('name', ''),
                        'scenarioName': scenario.get('name', '')
                    })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Search all people (walk-ins)
    # -----------------------------------------------------------------
    elif action == 'search_all_people':
        try:
            search_term = str(Data.search_term) if hasattr(Data, 'search_term') else ''
            safe_term = search_term.replace("'", "''")

            sql = """
                SELECT TOP 30 p.PeopleId,
                       ISNULL(p.Name2, '') as Name2,
                       ISNULL(p.FirstName, '') as FirstName,
                       ISNULL(p.LastName, '') as LastName,
                       ISNULL(p.NickName, '') as NickName,
                       p.Age,
                       ISNULL(g.Description, '') as Gender
                FROM People p
                LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                WHERE p.IsDeceased = 0
                  AND (p.LastName LIKE '%{0}%'
                       OR p.FirstName LIKE '%{0}%'
                       OR p.NickName LIKE '%{0}%'
                       OR p.Name2 LIKE '%{0}%')
                ORDER BY p.LastName, p.FirstName
            """.format(safe_term)

            results = q.QuerySql(sql)
            people = []
            for r in results:
                try:
                    age_val = r.Age
                except:
                    age_val = None
                people.append({
                    'peopleId': r.PeopleId,
                    'name': str(r.Name2),
                    'firstName': str(r.FirstName),
                    'lastName': str(r.LastName),
                    'nickName': str(r.NickName),
                    'age': str(age_val) if age_val else '',
                    'gender': str(r.Gender)
                })
            print json.dumps({'success': True, 'people': people})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Get person detail with family, emergency, and reg data
    # -----------------------------------------------------------------
    elif action == 'get_person_detail':
        debug_step = 'init'
        try:
            debug_step = 'parse_people_id'
            people_id = int(str(Data.people_id)) if hasattr(Data, 'people_id') else 0

            debug_step = 'parse_source_org_id'
            source_org_id = 0
            try:
                if hasattr(Data, 'source_org_id'):
                    soid = str(Data.source_org_id)
                    if soid and soid != '0' and soid != '':
                        source_org_id = int(soid)
            except:
                source_org_id = 0

            debug_step = 'person_query'
            sql = """
                SELECT p.PeopleId,
                       ISNULL(p.Name2, '') as Name2,
                       ISNULL(p.FirstName, '') as FirstName,
                       ISNULL(p.LastName, '') as LastName,
                       ISNULL(p.NickName, '') as NickName,
                       p.Age, p.FamilyId,
                       ISNULL(g.Description, '') as Gender
                FROM People p
                LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                WHERE p.PeopleId = {0}
            """.format(people_id)
            result = q.QuerySql(sql)

            person = None
            for row in result:
                person = row
                break

            if not person:
                print json.dumps({'success': False, 'message': 'Person not found'})
            else:
                debug_step = 'read_person_name'
                p_name = str(person.Name2)

                debug_step = 'read_person_first'
                p_first = str(person.FirstName)

                debug_step = 'read_person_last'
                p_last = str(person.LastName)

                debug_step = 'read_person_gender'
                p_gender = str(person.Gender)

                debug_step = 'read_person_age'
                age_val = ''
                try:
                    if person.Age is not None:
                        age_val = str(person.Age)
                except:
                    age_val = ''

                debug_step = 'read_person_famid'
                fam_id = None
                try:
                    fam_id = person.FamilyId
                    if fam_id is not None:
                        fam_id = int(fam_id)
                except:
                    fam_id = None

                debug_step = 'build_person_data'
                person_data = {
                    'peopleId': people_id,
                    'name': p_name,
                    'firstName': p_first,
                    'lastName': p_last,
                    'age': age_val,
                    'gender': p_gender,
                    'familyId': fam_id
                }

                # RecReg (emergency/medical) data
                debug_step = 'emergency_init'
                emergency = {
                    'contactName': '',
                    'contactPhone': '',
                    'medical': '',
                    'allergies': '',
                    'doctor': '',
                    'doctorPhone': '',
                    'insurance': '',
                    'policy': '',
                    'comments': '',
                    'otcMeds': []
                }
                try:
                    debug_step = 'recreg_query'
                    rr_sql = """
                        SELECT TOP 1
                               ISNULL(rr.emcontact, '') as emcontact,
                               ISNULL(rr.emphone, '') as emphone,
                               ISNULL(rr.MedAllergy, '') as MedAllergy,
                               ISNULL(rr.MedicalDescription, '') as MedicalDescription,
                               ISNULL(rr.doctor, '') as doctor,
                               ISNULL(rr.docphone, '') as docphone,
                               ISNULL(rr.insurance, '') as insurance,
                               ISNULL(rr.policy, '') as policy,
                               ISNULL(rr.Comments, '') as rrComments,
                               ISNULL(rr.Tylenol, 0) as Tylenol,
                               ISNULL(rr.Advil, 0) as Advil,
                               ISNULL(rr.Maalox, 0) as Maalox,
                               ISNULL(rr.Robitussin, 0) as Robitussin
                        FROM RecReg rr
                        WHERE rr.PeopleId = {0}
                    """.format(people_id)
                    rr_result = q.QuerySql(rr_sql)
                    rr = None
                    for row in rr_result:
                        rr = row
                        break
                    if rr:
                        debug_step = 'recreg_read'

                        def rr_str(field):
                            try:
                                v = getattr(rr, field, None)
                                if v is None:
                                    return ''
                                return str(v)
                            except:
                                return ''

                        # OTC meds: try sources in order, dedupe case-insensitive while preserving
                        # display order. PersonMedication is TouchPoint's current canonical store
                        # (populated by newer registration forms), RegAnswer is the in-form answer
                        # for older per-org questions, RecReg booleans are the legacy fallback.
                        otc_meds = []
                        otc_seen = set()

                        def _add_otc(name):
                            n = (name or '').strip().strip('"').strip("'").strip()
                            if not n:
                                return
                            k = n.lower()
                            if k in otc_seen:
                                return
                            otc_seen.add(k)
                            otc_meds.append(n)

                        # Source 1 (preferred): PersonMedication -> lookup.Medication
                        try:
                            debug_step = 'otc_personmedication'
                            pm_sql = """
                                SELECT med.Description
                                FROM dbo.PersonMedication pm
                                LEFT JOIN lookup.Medication med ON med.Id = pm.MedicationId
                                WHERE pm.PeopleId = {0}
                                  AND med.Description IS NOT NULL
                            """.format(people_id)
                            for row in q.QuerySql(pm_sql):
                                _add_otc(str(row.Description))
                        except:
                            pass

                        # Source 2: RegAnswer matched by question label (covers both original
                        # hardcoded GUID and newer per-org variants of the OTC question)
                        try:
                            debug_step = 'otc_reganswer'
                            otc_sql = """
                                SELECT TOP 1 ra.AnswerValue
                                FROM RegAnswer ra
                                INNER JOIN RegPeople rp ON rp.RegPeopleId = ra.RegPeopleId
                                INNER JOIN RegQuestion rq ON rq.RegQuestionId = ra.RegQuestionId
                                WHERE rp.PeopleId = {0}
                                  AND (
                                        ra.RegQuestionId = '8A9F1199-6F1C-480C-8A2D-146EEEAE55B8'
                                     OR rq.Label LIKE '%Approved%over-the-counter%'
                                     OR rq.Label LIKE '%Approved Over-the-Counter%'
                                     OR rq.Label LIKE '%OTC medication%'
                                  )
                                  AND ra.AnswerValue IS NOT NULL
                                  AND LTRIM(RTRIM(CAST(ra.AnswerValue AS NVARCHAR(MAX)))) <> ''
                                ORDER BY rp.RegPeopleId DESC
                            """.format(people_id)
                            otc_row = None
                            for row in q.QuerySql(otc_sql):
                                otc_row = row
                                break
                            if otc_row and otc_row.AnswerValue:
                                answer_value = str(otc_row.AnswerValue)
                                if answer_value.startswith('[') and answer_value.endswith(']'):
                                    answer_value = answer_value[1:-1]
                                for med in answer_value.split(','):
                                    _add_otc(med)
                        except:
                            pass

                        # Source 3 (legacy fallback): RecReg boolean flags
                        try:
                            if rr.Tylenol: _add_otc('Tylenol')
                        except: pass
                        try:
                            if rr.Advil: _add_otc('Advil')
                        except: pass
                        try:
                            if rr.Maalox: _add_otc('Maalox')
                        except: pass
                        try:
                            if rr.Robitussin: _add_otc('Robitussin')
                        except: pass

                        emergency = {
                            'contactName': rr_str('emcontact'),
                            'contactPhone': rr_str('emphone'),
                            'medical': rr_str('MedicalDescription'),
                            'allergies': rr_str('MedAllergy'),
                            'doctor': rr_str('doctor'),
                            'doctorPhone': rr_str('docphone'),
                            'insurance': rr_str('insurance'),
                            'policy': rr_str('policy'),
                            'comments': rr_str('rrComments'),
                            'otcMeds': otc_meds
                        }
                except:
                    pass

                # Registration answers from RegQuestion/RegAnswer
                debug_step = 'reg_answers'
                reg_answers = []
                if source_org_id:
                    try:
                        answers_sql = """
                            SELECT rq.Label as Question, rq.[Order] as QOrder, ra.AnswerValue as Answer
                            FROM Registration r WITH (NOLOCK)
                            JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
                            LEFT JOIN RegAnswer ra WITH (NOLOCK) ON rp.RegPeopleId = ra.RegPeopleId
                            LEFT JOIN RegQuestion rq WITH (NOLOCK) ON ra.RegQuestionId = rq.RegQuestionId
                            WHERE r.OrganizationId = {0}
                              AND rp.PeopleId = {1}
                              AND rq.Label IS NOT NULL
                            ORDER BY rq.[Order]
                        """.format(source_org_id, people_id)
                        answers_result = q.QuerySql(answers_sql)
                        if answers_result:
                            for ar in answers_result:
                                q_text = ''
                                a_text = ''
                                try:
                                    q_text = str(ar.Question) if ar.Question else ''
                                except: pass
                                try:
                                    a_text = str(ar.Answer) if ar.Answer else ''
                                except: pass
                                if q_text:
                                    reg_answers.append({'question': q_text, 'answer': a_text})
                    except:
                        pass

                    # RegistrationData XML fallback if no RegAnswer data
                    if not reg_answers:
                        try:
                            rd_sql = """
                                SELECT TOP 10 CAST(rd.Data AS NVARCHAR(MAX)) as XmlData
                                FROM RegistrationData rd WITH (NOLOCK)
                                WHERE rd.OrganizationId = {0}
                                  AND rd.completed = 1
                                ORDER BY rd.Stamp DESC
                            """.format(source_org_id)
                            rd_result = q.QuerySql(rd_sql)
                            if rd_result:
                                for row in rd_result:
                                    try:
                                        xml_data = row.XmlData
                                        if not xml_data:
                                            continue
                                        xml_data = str(xml_data)
                                    except:
                                        continue
                                    people_id_pattern = '<PeopleId>{0}</PeopleId>'.format(people_id)
                                    if people_id_pattern not in xml_data:
                                        continue

                                    person_pattern = r'<OnlineRegPersonModel[^>]*>.*?<PeopleId>{0}</PeopleId>.*?</OnlineRegPersonModel>'.format(people_id)
                                    person_match = re.search(person_pattern, xml_data, re.DOTALL)
                                    if not person_match:
                                        continue

                                    person_xml = person_match.group(0)
                                    seen_questions = set()

                                    extra_pattern = r'<ExtraQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</ExtraQuestion>'
                                    for match in re.finditer(extra_pattern, person_xml):
                                        qt = match.group(1)
                                        at = match.group(2).strip() if match.group(2) else ''
                                        if qt and qt not in seen_questions:
                                            reg_answers.append({'question': qt, 'answer': at})
                                            seen_questions.add(qt)

                                    text_pattern = r'<Text[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</Text>'
                                    for match in re.finditer(text_pattern, person_xml):
                                        qt = match.group(1)
                                        at = match.group(2).strip() if match.group(2) else ''
                                        if qt and qt not in seen_questions:
                                            reg_answers.append({'question': qt, 'answer': at})
                                            seen_questions.add(qt)

                                    yn_pattern = r'<YesNoQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</YesNoQuestion>'
                                    for match in re.finditer(yn_pattern, person_xml):
                                        qt = match.group(1)
                                        av = match.group(2).strip() if match.group(2) else ''
                                        if qt and qt not in seen_questions:
                                            at = 'Yes' if av == 'True' else 'No' if av == 'False' else av
                                            reg_answers.append({'question': qt, 'answer': at})
                                            seen_questions.add(qt)

                                    if reg_answers:
                                        break
                        except:
                            pass

                # Family members
                debug_step = 'family_query'
                family = []
                if fam_id:
                    try:
                        fam_sql = """
                            SELECT p.PeopleId,
                                   ISNULL(p.Name2, '') as Name2,
                                   p.Age,
                                   ISNULL(g.Description, '') as Gender,
                                   ISNULL(fp.Description, '') as FamilyPosition
                            FROM People p
                            LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                            LEFT JOIN lookup.FamilyPosition fp ON p.PositionInFamilyId = fp.Id
                            WHERE p.FamilyId = {0} AND p.PeopleId != {1}
                            ORDER BY p.PositionInFamilyId, p.Age DESC
                        """.format(fam_id, people_id)
                        fam_results = q.QuerySql(fam_sql)
                        for f in fam_results:
                            f_age = None
                            f_name = ''
                            f_gender = ''
                            f_position = ''
                            try:
                                f_age = f.Age
                            except: pass
                            try:
                                f_name = str(f.Name2) if f.Name2 else ''
                            except: pass
                            try:
                                f_gender = str(f.Gender) if f.Gender else ''
                            except: pass
                            try:
                                f_position = str(f.FamilyPosition) if f.FamilyPosition else ''
                            except: pass
                            family.append({
                                'peopleId': f.PeopleId,
                                'name': f_name,
                                'age': str(f_age) if f_age else '',
                                'gender': f_gender,
                                'position': f_position
                            })
                    except:
                        pass

                debug_step = 'response'
                print json.dumps({
                    'success': True,
                    'person': person_data,
                    'emergency': emergency,
                    'regAnswers': reg_answers,
                    'family': family
                })
        except Exception as e:
            print json.dumps({'success': False, 'message': 'Error at ' + debug_step + ': ' + str(e)})

    # -----------------------------------------------------------------
    # STATION: Get fresh destination counts (polling)
    # -----------------------------------------------------------------
    elif action == 'get_destination_counts':
        try:
            org_ids_str = str(Data.org_ids) if hasattr(Data, 'org_ids') else ''
            org_ids = [int(x.strip()) for x in org_ids_str.split(',') if x.strip()]

            counts = {}
            sg_counts = {}
            for oid in org_ids:
                cnt_sql = "SELECT Count(om.PeopleId) as cnt FROM OrganizationMembers om WHERE om.OrganizationId = {0}".format(oid)
                cnt_row = q.QuerySqlTop1(cnt_sql)
                counts[str(oid)] = cnt_row.cnt if cnt_row and hasattr(cnt_row, 'cnt') else 0
                # Subgroup counts
                sg_sql = "SELECT sg.Name, COUNT(DISTINCT ommt.PeopleId) as cnt FROM MemberTags sg LEFT JOIN OrgMemMemTags ommt ON sg.Id = ommt.MemberTagId WHERE sg.OrgId = {0} GROUP BY sg.Id, sg.Name".format(oid)
                org_sg = {}
                for sgr in q.QuerySql(sg_sql):
                    org_sg[sgr.Name] = sgr.cnt or 0
                if org_sg:
                    sg_counts[str(oid)] = org_sg

            print json.dumps({'success': True, 'counts': counts, 'subgroupCounts': sg_counts})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Process assignment (core action)
    # -----------------------------------------------------------------
    elif action == 'process_assignment':
        try:
            people_id = int(str(Data.people_id)) if hasattr(Data, 'people_id') else 0
            dest_org_id = int(str(Data.dest_org_id)) if hasattr(Data, 'dest_org_id') else 0
            dest_subgroup = str(Data.dest_subgroup) if hasattr(Data, 'dest_subgroup') else ''
            capacity = int(str(Data.capacity)) if hasattr(Data, 'capacity') else 0
            cap_type = str(Data.cap_type) if hasattr(Data, 'cap_type') else 'soft'
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            station_id = str(Data.station_id) if hasattr(Data, 'station_id') else ''
            station_name = str(Data.station_name) if hasattr(Data, 'station_name') else ''
            volunteer_name = str(Data.volunteer_name) if hasattr(Data, 'volunteer_name') else ''
            dest_display_name = str(Data.dest_display_name) if hasattr(Data, 'dest_display_name') else ''
            dest_org_name = str(Data.dest_org_name) if hasattr(Data, 'dest_org_name') else ''

            # Fresh count check
            cnt_sql = "SELECT Count(om.PeopleId) as cnt FROM OrganizationMembers om WHERE om.OrganizationId = {0}".format(dest_org_id)
            cnt_row = q.QuerySqlTop1(cnt_sql)
            current_count = cnt_row.cnt if cnt_row and hasattr(cnt_row, 'cnt') else 0

            over_cap = False
            if capacity > 0 and current_count >= capacity:
                if cap_type == 'hard':
                    print json.dumps({
                        'success': False,
                        'message': 'Hard cap reached ({0}/{1}). Cannot assign.'.format(current_count, capacity),
                        'currentCount': current_count,
                        'blocked': True
                    })
                else:
                    over_cap = True

            if not (capacity > 0 and current_count >= capacity and cap_type == 'hard'):
                person = model.GetPerson(people_id)
                if not person:
                    print json.dumps({'success': False, 'message': 'Person not found'})
                else:
                    if not model.InOrg(people_id, dest_org_id):
                        model.JoinOrg(dest_org_id, person)

                    if dest_subgroup:
                        # Support multiple subgroups separated by pipe
                        sg_list = dest_subgroup.split('|') if '|' in dest_subgroup else [dest_subgroup]
                        for sg in sg_list:
                            sg = sg.strip()
                            if sg:
                                model.AddSubGroup(people_id, dest_org_id, sg)

                    # Log the assignment
                    log_key = "DayOfRegistration_Log_" + scenario_id
                    log_raw = model.TextContent(log_key) or ''
                    log_data = json.loads(log_raw) if log_raw else {'entries': []}
                    log_data['entries'].append({
                        'peopleId': people_id,
                        'personName': person.Name2 or '',
                        'destOrgId': dest_org_id,
                        'destOrgName': dest_org_name,
                        'destDisplayName': dest_display_name,
                        'destSubgroup': dest_subgroup,
                        'stationId': station_id,
                        'stationName': station_name,
                        'volunteerName': volunteer_name,
                        'processedAt': datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        'dr_action': 'assign'
                    })
                    model.WriteContentText(log_key, json.dumps(log_data), "")

                    # Re-query actual count after join
                    post_cnt_sql = "SELECT Count(om.PeopleId) as cnt FROM OrganizationMembers om WHERE om.OrganizationId = {0}".format(dest_org_id)
                    post_cnt_row = q.QuerySqlTop1(post_cnt_sql)
                    new_count = post_cnt_row.cnt if post_cnt_row and hasattr(post_cnt_row, 'cnt') else current_count + 1
                    # Get updated subgroup counts
                    post_sg_counts = {}
                    sg_sql = "SELECT sg.Name, COUNT(DISTINCT ommt.PeopleId) as cnt FROM MemberTags sg LEFT JOIN OrgMemMemTags ommt ON sg.Id = ommt.MemberTagId WHERE sg.OrgId = {0} GROUP BY sg.Id, sg.Name".format(dest_org_id)
                    for sgr in q.QuerySql(sg_sql):
                        post_sg_counts[sgr.Name] = sgr.cnt or 0
                    print json.dumps({
                        'success': True,
                        'message': 'Assigned ' + (person.Name2 or '') + ' to ' + dest_display_name,
                        'currentCount': new_count,
                        'subgroupCounts': post_sg_counts,
                        'overCap': over_cap,
                        'personName': person.Name2 or ''
                    })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Undo assignment
    # -----------------------------------------------------------------
    elif action == 'undo_assignment':
        try:
            people_id = int(str(Data.people_id)) if hasattr(Data, 'people_id') else 0
            dest_org_id = int(str(Data.dest_org_id)) if hasattr(Data, 'dest_org_id') else 0
            dest_subgroup = str(Data.dest_subgroup) if hasattr(Data, 'dest_subgroup') else ''
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            station_id = str(Data.station_id) if hasattr(Data, 'station_id') else ''
            station_name = str(Data.station_name) if hasattr(Data, 'station_name') else ''
            volunteer_name = str(Data.volunteer_name) if hasattr(Data, 'volunteer_name') else ''
            dest_display_name = str(Data.dest_display_name) if hasattr(Data, 'dest_display_name') else ''
            dest_org_name = str(Data.dest_org_name) if hasattr(Data, 'dest_org_name') else ''

            person = model.GetPerson(people_id)
            person_name = person.Name2 if person else str(people_id)

            if dest_subgroup:
                sg_list = dest_subgroup.split('|') if '|' in dest_subgroup else [dest_subgroup]
                for sg in sg_list:
                    sg = sg.strip()
                    if sg:
                        model.RemoveSubGroup(people_id, dest_org_id, sg)

            if person:
                model.DropOrgMember(people_id, dest_org_id)

            # Log the undo
            log_key = "DayOfRegistration_Log_" + scenario_id
            log_raw = model.TextContent(log_key) or ''
            log_data = json.loads(log_raw) if log_raw else {'entries': []}
            log_data['entries'].append({
                'peopleId': people_id,
                'personName': person_name,
                'destOrgId': dest_org_id,
                'destOrgName': dest_org_name,
                'destDisplayName': dest_display_name,
                'destSubgroup': dest_subgroup,
                'stationId': station_id,
                'stationName': station_name,
                'volunteerName': volunteer_name,
                'processedAt': datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                'dr_action': 'undo'
            })
            model.WriteContentText(log_key, json.dumps(log_data), "")

            cnt_sql = "SELECT Count(om.PeopleId) as cnt FROM OrganizationMembers om WHERE om.OrganizationId = {0}".format(dest_org_id)
            cnt_row = q.QuerySqlTop1(cnt_sql)
            new_count = cnt_row.cnt if cnt_row and hasattr(cnt_row, 'cnt') else 0
            # Get updated subgroup counts
            undo_sg_counts = {}
            sg_sql = "SELECT sg.Name, COUNT(DISTINCT ommt.PeopleId) as cnt FROM MemberTags sg LEFT JOIN OrgMemMemTags ommt ON sg.Id = ommt.MemberTagId WHERE sg.OrgId = {0} GROUP BY sg.Id, sg.Name".format(dest_org_id)
            for sgr in q.QuerySql(sg_sql):
                undo_sg_counts[sgr.Name] = sgr.cnt or 0

            print json.dumps({
                'success': True,
                'message': 'Removed ' + person_name + ' from ' + dest_display_name,
                'currentCount': new_count,
                'subgroupCounts': undo_sg_counts
            })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Find friend (search across all destinations)
    # -----------------------------------------------------------------
    elif action == 'find_friend':
        try:
            search_term = str(Data.search_term) if hasattr(Data, 'search_term') else ''
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''

            safe_term = search_term.replace("'", "''")

            # Load scenario to get filter settings and destination map
            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}
            scenario = None
            for s in data.get('scenarios', []):
                if s.get('id') == scenario_id:
                    scenario = s
                    break

            # Get friend finder program/division filter from scenario
            ff_program_id = scenario.get('friendFinderProgramId', '') if scenario else ''
            ff_division_id = scenario.get('friendFinderDivisionId', '') if scenario else ''

            # Build destination org map. Fall back to orgName when displayName is blank
            # (Bulk Add leaves displayName empty by design; the rest of the UI defaults
            # to orgName at render time, so we mirror that here for consistency).
            dest_org_map = {}
            if scenario:
                for station in scenario.get('stations', []):
                    for dest in station.get('destinations', []):
                        dn = dest.get('displayName') or dest.get('orgName') or ''
                        dest_org_map[dest.get('orgId')] = {
                            'displayName': dn,
                            'stationName': station.get('name', '')
                        }

            # Search people - optionally filtered to those in orgs within program/division.
            # Use OrganizationStructure so secondary div/prog memberships count too (an
            # involvement can sit in a Primary AND Secondary division; Organizations.DivisionId
            # only holds the primary, which is why Find-a-Friend was missing rooms whose primary
            # division differed from the configured friend-finder filter).
            if ff_program_id or ff_division_id:
                filter_clauses = ["os.OrgStatus = 'Active'"]
                if ff_program_id:
                    filter_clauses.append("os.ProgId = {0}".format(int(ff_program_id)))
                if ff_division_id:
                    filter_clauses.append("os.DivId = {0}".format(int(ff_division_id)))

                sql = """
                    SELECT TOP 20 p.PeopleId,
                           ISNULL(p.Name2, '') as Name2,
                           p.Age,
                           ISNULL(g.Description, '') as Gender
                    FROM People p
                    LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                    WHERE p.IsDeceased = 0
                      AND (p.LastName LIKE '%{0}%'
                           OR p.FirstName LIKE '%{0}%'
                           OR p.NickName LIKE '%{0}%'
                           OR p.Name2 LIKE '%{0}%')
                      AND p.PeopleId IN (
                          SELECT DISTINCT om.PeopleId
                          FROM OrganizationMembers om
                          JOIN dbo.OrganizationStructure os ON os.OrgId = om.OrganizationId
                          WHERE {1}
                      )
                    ORDER BY p.LastName, p.FirstName
                """.format(safe_term, " AND ".join(filter_clauses))
            else:
                sql = """
                    SELECT TOP 20 p.PeopleId,
                           ISNULL(p.Name2, '') as Name2,
                           p.Age,
                           ISNULL(g.Description, '') as Gender
                    FROM People p
                    LEFT JOIN lookup.Gender g ON p.GenderId = g.Id
                    WHERE p.IsDeceased = 0
                      AND (p.LastName LIKE '%{0}%'
                           OR p.FirstName LIKE '%{0}%'
                           OR p.NickName LIKE '%{0}%'
                           OR p.Name2 LIKE '%{0}%')
                    ORDER BY p.LastName, p.FirstName
                """.format(safe_term)

            results = q.QuerySql(sql)
            people_ids = []
            people_list = []
            for r in results:
                try:
                    age_val = r.Age
                except:
                    age_val = None
                people_ids.append(r.PeopleId)
                people_list.append({
                    'peopleId': r.PeopleId,
                    'name': str(r.Name2),
                    'age': str(age_val) if age_val else '',
                    'gender': str(r.Gender)
                })

            # Batch query: get filtered involvements for all matched people.
            # OrganizationStructure can return the same OrgId twice when an involvement
            # belongs to both a primary and secondary division within scope, so dedupe.
            inv_map = {}  # peopleId -> [{orgId, orgName}]
            if people_ids and (ff_program_id or ff_division_id):
                inv_filter = ["os.OrgStatus = 'Active'"]
                if ff_program_id:
                    inv_filter.append("os.ProgId = {0}".format(int(ff_program_id)))
                if ff_division_id:
                    inv_filter.append("os.DivId = {0}".format(int(ff_division_id)))

                inv_sql = """
                    SELECT DISTINCT om.PeopleId, os.OrgId AS OrganizationId,
                           ISNULL(os.Organization, '') as OrganizationName
                    FROM OrganizationMembers om
                    JOIN dbo.OrganizationStructure os ON os.OrgId = om.OrganizationId
                    WHERE om.PeopleId IN ({0})
                      AND {1}
                    ORDER BY om.PeopleId, OrganizationName
                """.format(','.join(str(pid) for pid in people_ids), " AND ".join(inv_filter))

                seen_inv = {}  # pid -> set of orgIds (manual dedupe; SELECT DISTINCT alone
                               # doesn't help if the row tuple differs only by which prog/div
                               # path matched, but we project just OrgId/Name above)
                for inv_r in q.QuerySql(inv_sql):
                    pid = inv_r.PeopleId
                    if pid not in inv_map:
                        inv_map[pid] = []
                        seen_inv[pid] = set()
                    if inv_r.OrganizationId in seen_inv[pid]:
                        continue
                    seen_inv[pid].add(inv_r.OrganizationId)
                    inv_map[pid].append({
                        'orgId': inv_r.OrganizationId,
                        'orgName': str(inv_r.OrganizationName)
                    })

            # Build results with destination checks and filtered involvements.
            # Strip any involvements that are already shown as destination badges so the
            # same room doesn't render twice (once green, once blue).
            found = []
            for p in people_list:
                pid = p['peopleId']
                person_dests = []
                dest_orgids_seen = set()
                for org_id, info in dest_org_map.items():
                    if model.InOrg(pid, org_id):
                        person_dests.append({
                            'displayName': info['displayName'],
                            'stationName': info['stationName']
                        })
                        dest_orgids_seen.add(org_id)
                p['destinations'] = person_dests
                p['involvements'] = [
                    inv for inv in inv_map.get(pid, [])
                    if inv.get('orgId') not in dest_orgids_seen
                ]
                found.append(p)

            print json.dumps({'success': True, 'results': found})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # STATION: Get sync status (which source people are processed)
    # -----------------------------------------------------------------
    elif action == 'get_sync_status':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''

            # Load scenario config to get all destination org IDs
            raw = model.TextContent("DayOfRegistration_Scenarios") or ''
            data = json.loads(raw) if raw else {'scenarios': []}
            scenario = None
            for s in data.get('scenarios', []):
                if s.get('id') == scenario_id:
                    scenario = s
                    break

            assigned = {}
            if scenario:
                dest_org_ids = []
                dest_display_map = {}
                for st in scenario.get('stations', []):
                    for d in st.get('destinations', []):
                        oid = d.get('orgId', 0)
                        if oid and oid not in dest_org_ids:
                            dest_org_ids.append(oid)
                            dest_display_map[oid] = d.get('displayName', '') or d.get('orgName', '')

                if dest_org_ids:
                    sync_sql = """
                        SELECT om.PeopleId, om.OrganizationId,
                               o.OrganizationName,
                               ISNULL(p.Name2, '') as Name2
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        JOIN People p ON om.PeopleId = p.PeopleId
                        WHERE om.OrganizationId IN ({0})
                    """.format(','.join(str(x) for x in dest_org_ids))
                    for row in q.QuerySql(sync_sql):
                        pid = row.PeopleId
                        oid = row.OrganizationId
                        assigned[str(pid)] = {
                            'destOrgId': oid,
                            'destDisplayName': dest_display_map.get(oid, str(row.OrganizationName or '')),
                            'destOrgName': str(row.OrganizationName or ''),
                            'personName': str(row.Name2 or '')
                        }

            print json.dumps({'success': True, 'assigned': assigned})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # ADMIN: Load registration questions for an organization
    # -----------------------------------------------------------------
    elif action == 'load_reg_questions':
        try:
            org_id = int(str(Data.org_id)) if hasattr(Data, 'org_id') else 0
            if not org_id:
                print json.dumps({'success': False, 'message': 'org_id required'})
            else:
                questions = []
                seen = set()
                order_counter = 0

                # Approach 1: RegQuestion/RegAnswer tables (newer registrations)
                try:
                    sql = """
                        SELECT DISTINCT rq.Label, rq.[Order]
                        FROM Registration r WITH (NOLOCK)
                        JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
                        JOIN RegAnswer ra WITH (NOLOCK) ON rp.RegPeopleId = ra.RegPeopleId
                        JOIN RegQuestion rq WITH (NOLOCK) ON ra.RegQuestionId = rq.RegQuestionId
                        WHERE r.OrganizationId = {0}
                          AND rq.Label IS NOT NULL
                        ORDER BY rq.[Order]
                    """.format(org_id)
                    results = q.QuerySql(sql)
                    for row in results:
                        label = str(row.Label) if row.Label else ''
                        if label and label not in seen:
                            seen.add(label)
                            questions.append({'label': label, 'order': row.Order or 0})
                            order_counter = max(order_counter, (row.Order or 0) + 1)
                except:
                    pass

                # Approach 2: RegistrationData XML fallback (older registrations)
                try:
                    rd_sql = """
                        SELECT TOP 10 CAST(rd.Data AS NVARCHAR(MAX)) as XmlData
                        FROM RegistrationData rd WITH (NOLOCK)
                        WHERE rd.OrganizationId = {0}
                          AND rd.completed = 1
                        ORDER BY rd.Stamp DESC
                    """.format(org_id)
                    rd_result = q.QuerySql(rd_sql)
                    if rd_result:
                        for row in rd_result:
                            try:
                                xml_data = str(row.XmlData) if row.XmlData else ''
                            except:
                                continue
                            if not xml_data:
                                continue
                            for pattern in [
                                r'<ExtraQuestion[^>]*\squestion="([^"]+)"',
                                r'<Text[^>]*\squestion="([^"]+)"',
                                r'<YesNoQuestion[^>]*\squestion="([^"]+)"'
                            ]:
                                for match in re.finditer(pattern, xml_data):
                                    label = match.group(1)
                                    if label and label not in seen:
                                        seen.add(label)
                                        questions.append({'label': label, 'order': order_counter})
                                        order_counter += 1
                except:
                    pass

                # Approach 3: RegQuestion linked to Registration directly (no registrants needed)
                try:
                    rq_sql = """
                        SELECT DISTINCT rq.Label, rq.[Order]
                        FROM Registration r WITH (NOLOCK)
                        JOIN RegQuestion rq WITH (NOLOCK) ON rq.RegistrationId = r.RegistrationId
                        WHERE r.OrganizationId = {0}
                          AND rq.Label IS NOT NULL
                        ORDER BY rq.[Order]
                    """.format(org_id)
                    rq_results = q.QuerySql(rq_sql)
                    for row in rq_results:
                        label = str(row.Label) if row.Label else ''
                        if label and label not in seen:
                            seen.add(label)
                            questions.append({'label': label, 'order': row.Order or order_counter})
                            order_counter = max(order_counter, (row.Order or 0) + 1)
                except:
                    pass

                # Approach 4: Parse Organizations.RegSetting config text directly (works with no registrants at all)
                try:
                    rs_sql = """
                        SELECT TOP 1 CAST(RegSetting AS NVARCHAR(MAX)) as RegSetting,
                               CAST(RegSettingXml AS NVARCHAR(MAX)) as RegSettingXml
                        FROM Organizations WITH (NOLOCK)
                        WHERE OrganizationId = {0}
                    """.format(org_id)
                    rs_results = q.QuerySql(rs_sql)

                    def _decode_entities(s):
                        if not s: return ''
                        return (s.replace('&lt;', '<').replace('&gt;', '>')
                                 .replace('&quot;', '"').replace('&apos;', "'")
                                 .replace('&#39;', "'").replace('&amp;', '&')).strip()

                    def _add_q(label):
                        label = _decode_entities(label)
                        if label and label not in seen:
                            seen.add(label)
                            return True
                        return False

                    for row in rs_results:
                        configs = []
                        try:
                            if row.RegSetting: configs.append(str(row.RegSetting))
                        except: pass
                        try:
                            if row.RegSettingXml: configs.append(str(row.RegSettingXml))
                        except: pass
                        for cfg in configs:
                            if not cfg:
                                continue
                            # XML format: <Question>text</Question>, <Label>text</Label>
                            for pattern in [
                                r'<Question>([^<]+)</Question>',
                                r'<Label>([^<]+)</Label>'
                            ]:
                                for match in re.finditer(pattern, cfg, re.IGNORECASE | re.DOTALL):
                                    label = match.group(1)
                                    if _add_q(label):
                                        questions.append({'label': _decode_entities(label), 'order': order_counter})
                                        order_counter += 1
                            # Built-in askers without custom labels - map to friendly names
                            builtin_map = [
                                (r'<AskAllergies\b', 'Allergies'),
                                (r'<AskEmContact\b', 'Emergency Contact'),
                                (r'<AskMedical\b', 'Medical Information'),
                                (r'<AskTylenolEtc\b', 'Tylenol/Medication Permissions'),
                                (r'<AskTshirtSize\b', 'T-Shirt Size'),
                                (r'<AskChurch\b', 'Church You Attend'),
                                (r'<AskCoaching\b', 'Coaching Preference'),
                                (r'<AskParents\b', 'Parents'),
                                (r'<AskRequest\b', 'Request')
                            ]
                            for pat, friendly in builtin_map:
                                if re.search(pat, cfg, re.IGNORECASE) and friendly not in seen:
                                    seen.add(friendly)
                                    questions.append({'label': friendly, 'order': order_counter})
                                    order_counter += 1
                            # Legacy fallbacks for older Python-style RegSetting
                            for pattern in [
                                r'\bQuestion\s*,\s*"([^"]+)"',
                                r'\bAskText\s*\(\s*"([^"]+)"'
                            ]:
                                for match in re.finditer(pattern, cfg, re.IGNORECASE):
                                    label = match.group(1)
                                    if _add_q(label):
                                        questions.append({'label': _decode_entities(label), 'order': order_counter})
                                        order_counter += 1
                except:
                    pass

                print json.dumps({'success': True, 'questions': questions})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # CLEAR ASSIGNMENT LOG: empty the per-scenario log file. Always safe
    # to run — only touches the audit log, doesn't change any people's
    # involvement membership.
    # -----------------------------------------------------------------
    elif action == 'clear_assignment_log':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            if not scenario_id:
                print json.dumps({'success': False, 'message': 'scenario_id required'})
            else:
                log_key = "DayOfRegistration_Log_" + scenario_id
                model.WriteContentText(log_key, json.dumps({'entries': []}), "")
                print json.dumps({'success': True, 'message': 'Assignment log cleared.'})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # RESET TEST ASSIGNMENTS - PREVIEW: list ONLY the people this script
    # added during testing (read from the audit log, with assigns canceled
    # by later undos already excluded). Does not show pre-existing room
    # members. Read-only.
    # -----------------------------------------------------------------
    elif action == 'reset_test_assignments_preview':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            if not scenario_id:
                print json.dumps({'success': False, 'message': 'scenario_id required'})
            else:
                raw = model.TextContent("DayOfRegistration_Scenarios") or ''
                data = json.loads(raw) if raw else {'scenarios': []}
                scenario = None
                for s in data.get('scenarios', []):
                    if s.get('id') == scenario_id:
                        scenario = s
                        break
                if not scenario:
                    print json.dumps({'success': False, 'message': 'Scenario not found'})
                else:
                    is_test_mode = bool(scenario.get('isTestMode'))

                    # Build display map AND active-orgs set from the current scenario config.
                    # Anything in the log pointing at an orgId not in this set is stale and
                    # gets dropped from the preview/reset.
                    dest_label_map = {}
                    active_org_ids = set()
                    for st in scenario.get('stations', []):
                        for d in st.get('destinations', []):
                            oid = d.get('orgId', 0)
                            if oid:
                                active_org_ids.add(oid)
                                dest_label_map[oid] = {
                                    'displayName': d.get('displayName') or d.get('orgName') or '',
                                    'stationName': st.get('name', '')
                                }

                    raw_pairs = _script_assigned_pairs(scenario_id)
                    log_total = len(raw_pairs)

                    # Filter 1: drop pairs whose orgId is no longer in this scenario
                    pairs_in_scope = [p for p in raw_pairs if p['orgId'] in active_org_ids]
                    skipped_stale_orgs = log_total - len(pairs_in_scope)

                    # Filter 2: verify each remaining person is still actually in that org
                    # (drops entries for people manually cleaned up since)
                    pairs = []
                    skipped_already_removed = 0
                    for p in pairs_in_scope:
                        try:
                            if model.InOrg(p['peopleId'], p['orgId']):
                                pairs.append(p)
                            else:
                                skipped_already_removed += 1
                        except:
                            skipped_already_removed += 1

                    # Per-room rollup
                    per_org = {}
                    for p in pairs:
                        oid = p['orgId']
                        if oid not in per_org:
                            label = dest_label_map.get(oid, {})
                            per_org[oid] = {
                                'orgId': oid,
                                'displayName': label.get('displayName') or p.get('displayName') or p.get('orgName') or '',
                                'orgName': p.get('orgName', ''),
                                'stationName': label.get('stationName') or p.get('stationName') or '',
                                'count': 0,
                                'people': []
                            }
                        per_org[oid]['count'] += 1
                        per_org[oid]['people'].append({
                            'peopleId': p['peopleId'],
                            'name': p['name']
                        })

                    print json.dumps({
                        'success': True,
                        'scenarioName': scenario.get('name', ''),
                        'isTestMode': is_test_mode,
                        'totalAssigned': len(pairs),
                        'logTotal': log_total,
                        'skippedStaleOrgs': skipped_stale_orgs,
                        'skippedAlreadyRemoved': skipped_already_removed,
                        'destinations': list(per_org.values())
                    })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # RESET TEST ASSIGNMENTS - EXECUTE: remove ONLY the people the
    # script logged as assigning AND who are still in scenario destinations
    # AND still actually members of the org. Pre-existing members of the
    # room are never touched. Snapshots first; gated by isTestMode + name.
    # -----------------------------------------------------------------
    elif action == 'reset_test_assignments':
        try:
            scenario_id = str(Data.scenario_id) if hasattr(Data, 'scenario_id') else ''
            confirm_name = str(Data.confirm_name) if hasattr(Data, 'confirm_name') else ''
            if not scenario_id:
                print json.dumps({'success': False, 'message': 'scenario_id required'})
            else:
                raw = model.TextContent("DayOfRegistration_Scenarios") or ''
                data = json.loads(raw) if raw else {'scenarios': []}
                scenario = None
                for s in data.get('scenarios', []):
                    if s.get('id') == scenario_id:
                        scenario = s
                        break
                if not scenario:
                    print json.dumps({'success': False, 'message': 'Scenario not found'})
                elif not bool(scenario.get('isTestMode')):
                    print json.dumps({'success': False, 'message': 'Reset blocked: scenario is not in Test Mode. Enable Test Mode in the scenario editor first.'})
                elif (confirm_name or '').strip() != (scenario.get('name') or '').strip():
                    print json.dumps({'success': False, 'message': 'Confirmation name does not match the scenario name. Reset cancelled.'})
                else:
                    # Apply the same two filters as the preview so the execute path
                    # can never act on stale log entries.
                    active_org_ids = set()
                    for st in scenario.get('stations', []):
                        for d in st.get('destinations', []):
                            oid = d.get('orgId', 0)
                            if oid:
                                active_org_ids.add(oid)
                    raw_pairs = _script_assigned_pairs(scenario_id)
                    pairs_in_scope = [p for p in raw_pairs if p['orgId'] in active_org_ids]
                    pairs = []
                    for p in pairs_in_scope:
                        try:
                            if model.InOrg(p['peopleId'], p['orgId']):
                                pairs.append(p)
                        except:
                            pass

                    # Snapshot the log + the affected pairs to a backup before destructive ops
                    log_key = "DayOfRegistration_Log_" + scenario_id
                    try:
                        log_raw = model.TextContent(log_key) or ''
                    except:
                        log_raw = ''
                    snapshot = {
                        'snapshotAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'scenarioId': scenario_id,
                        'scenarioName': scenario.get('name', ''),
                        'pairs': pairs,
                        'log': log_raw
                    }
                    backup_key = "DayOfRegistration_Backup_" + scenario_id + "_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                    model.WriteContentText(backup_key, json.dumps(snapshot), "")

                    # Remove ONLY the pairs the script assigned. DropOrgMember also clears
                    # the person's OrgMemMemTags (subgroup tags) for that org.
                    removed_count = 0
                    errors = []
                    for p in pairs:
                        try:
                            model.DropOrgMember(p['peopleId'], p['orgId'])
                            removed_count += 1
                        except Exception as re:
                            errors.append('PeopleId={0} orgId={1}: {2}'.format(p['peopleId'], p['orgId'], str(re)))

                    # Clear the assignment log
                    model.WriteContentText(log_key, json.dumps({'entries': []}), "")

                    print json.dumps({
                        'success': True,
                        'message': 'Reset complete. Removed {0} script-assigned person(s). Backup saved as "{1}".'.format(removed_count, backup_key),
                        'removed': removed_count,
                        'backupKey': backup_key,
                        'errors': errors
                    })
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -----------------------------------------------------------------
    # APPLY UPDATE: fetch latest script from DisplayCache and overwrite
    # the installed Python content slot. Triggered by Update Now banner.
    # -----------------------------------------------------------------
    elif action == 'apply_update':
        new_code = ''
        try:
            fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
            new_code = str(model.RestGet(fetch_url, {}))
        except Exception as fe:
            print json.dumps({'success': False, 'message': 'Failed to fetch update: ' + str(fe)})
        else:
            if not new_code or len(new_code) < 200:
                print json.dumps({'success': False, 'message': 'Invalid or empty script code received'})
            else:
                target_name = get_script_name() or DC_SCRIPT_ID
                try:
                    model.WriteContentPython(target_name, new_code)
                    print json.dumps({'success': True, 'message': 'Updated ' + target_name + '. Reload the page.'})
                except Exception as we:
                    print json.dumps({'success': False, 'message': 'Write failed: ' + str(we)})

    # -----------------------------------------------------------------
    # Unknown action
    # -----------------------------------------------------------------
    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# =====================================================================
# HTML OUTPUT (GET)
# =====================================================================
else:
    html = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <script>
    if (window.location.pathname.indexOf('/PyScript/') > -1) {
        window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
    }
    </script>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
/* ================================================================
   Day of Registration - Scoped CSS (dr- prefix)
   ================================================================ */
.dr-root {
    --dr-primary: #2563eb;
    --dr-primary-light: #dbeafe;
    --dr-secondary: #059669;
    --dr-secondary-light: #d1fae5;
    --dr-warning: #d97706;
    --dr-warning-light: #fef3c7;
    --dr-danger: #dc2626;
    --dr-danger-light: #fee2e2;
    --dr-dark: #1e293b;
    --dr-gray: #64748b;
    --dr-light-bg: #f8fafc;
    --dr-border: #e2e8f0;
    --dr-radius: 12px;
    --dr-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
    --dr-shadow-lg: 0 10px 15px rgba(0,0,0,0.1), 0 4px 6px rgba(0,0,0,0.05);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: var(--dr-dark);
    -webkit-text-size-adjust: 100%;
}

/* Destination color palette */
.dr-dest-0 { --dr-dest-color: #3b82f6; --dr-dest-bg: #dbeafe; }
.dr-dest-1 { --dr-dest-color: #059669; --dr-dest-bg: #d1fae5; }
.dr-dest-2 { --dr-dest-color: #d97706; --dr-dest-bg: #fef3c7; }
.dr-dest-3 { --dr-dest-color: #7c3aed; --dr-dest-bg: #ede9fe; }
.dr-dest-4 { --dr-dest-color: #db2777; --dr-dest-bg: #fce7f3; }
.dr-dest-5 { --dr-dest-color: #0891b2; --dr-dest-bg: #cffafe; }
.dr-dest-6 { --dr-dest-color: #ea580c; --dr-dest-bg: #ffedd5; }
.dr-dest-7 { --dr-dest-color: #4f46e5; --dr-dest-bg: #e0e7ff; }

.dr-container {
    max-width: 1400px;
    margin: 0 auto;
    padding: 12px;
}

/* ---- Landing Page ---- */
.dr-landing {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 70vh;
    gap: 24px;
}

.dr-landing-title {
    font-size: 32px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 16px;
}

.dr-landing-subtitle {
    font-size: 18px;
    color: var(--dr-gray);
    text-align: center;
    margin-bottom: 32px;
}

.dr-landing-buttons {
    display: flex;
    gap: 24px;
    flex-wrap: wrap;
    justify-content: center;
}

.dr-landing-btn {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    width: 260px;
    height: 200px;
    border-radius: var(--dr-radius);
    border: 2px solid var(--dr-border);
    background: white;
    cursor: pointer;
    transition: all 0.2s;
    text-decoration: none;
    color: var(--dr-dark);
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
    box-shadow: var(--dr-shadow);
}

.dr-landing-btn:active {
    transform: scale(0.97);
    box-shadow: var(--dr-shadow-lg);
}

.dr-landing-btn i {
    font-size: 48px;
    margin-bottom: 16px;
}

.dr-landing-btn span {
    font-size: 22px;
    font-weight: 600;
}

.dr-landing-btn.dr-admin-btn { border-color: var(--dr-primary); }
.dr-landing-btn.dr-admin-btn i { color: var(--dr-primary); }
.dr-landing-btn.dr-station-btn { border-color: var(--dr-secondary); }
.dr-landing-btn.dr-station-btn i { color: var(--dr-secondary); }

/* ---- Common UI ---- */
.dr-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 12px 24px;
    border-radius: 8px;
    border: none;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
    min-height: 48px;
}

.dr-btn:active { transform: scale(0.97); }
.dr-btn-primary { background: var(--dr-primary); color: white; }
.dr-btn-success { background: var(--dr-secondary); color: white; }
.dr-btn-warning { background: var(--dr-warning); color: white; }
.dr-btn-danger { background: var(--dr-danger); color: white; }
.dr-btn-outline { background: white; color: var(--dr-dark); border: 2px solid var(--dr-border); }
.dr-btn-sm { padding: 8px 16px; font-size: 14px; min-height: 40px; }
.dr-btn-lg { padding: 16px 32px; font-size: 20px; min-height: 60px; }
.dr-btn-block { width: 100%; }
.dr-btn:disabled { opacity: 0.5; cursor: not-allowed; }

.dr-input {
    width: 100%;
    padding: 12px 16px;
    border: 2px solid var(--dr-border);
    border-radius: 8px;
    font-size: 16px;
    outline: none;
    transition: border-color 0.15s;
    box-sizing: border-box;
    min-height: 48px;
}

.dr-input:focus { border-color: var(--dr-primary); }

.dr-select {
    width: 100%;
    padding: 12px 16px;
    border: 2px solid var(--dr-border);
    border-radius: 8px;
    font-size: 16px;
    background: white;
    min-height: 48px;
    box-sizing: border-box;
}

.dr-label {
    display: block;
    font-size: 14px;
    font-weight: 600;
    color: var(--dr-gray);
    margin-bottom: 6px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.dr-panel {
    background: white;
    border-radius: var(--dr-radius);
    border: 1px solid var(--dr-border);
    box-shadow: var(--dr-shadow);
    margin-bottom: 16px;
    overflow: hidden;
}

.dr-panel-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 16px 20px;
    background: var(--dr-light-bg);
    border-bottom: 1px solid var(--dr-border);
    font-weight: 600;
    font-size: 16px;
}

.dr-panel-body { padding: 20px; }

.dr-flex { display: flex; }
.dr-flex-wrap { flex-wrap: wrap; }
.dr-items-center { align-items: center; }
.dr-justify-between { justify-content: space-between; }
.dr-gap-8 { gap: 8px; }
.dr-gap-12 { gap: 12px; }
.dr-gap-16 { gap: 16px; }
.dr-mb-8 { margin-bottom: 8px; }
.dr-mb-12 { margin-bottom: 12px; }
.dr-mb-16 { margin-bottom: 16px; }
.dr-mb-24 { margin-bottom: 24px; }
.dr-mt-12 { margin-top: 12px; }
.dr-mt-16 { margin-top: 16px; }
.dr-text-center { text-align: center; }
.dr-text-muted { color: var(--dr-gray); }
.dr-text-sm { font-size: 14px; }
.dr-text-lg { font-size: 20px; }
.dr-fw-bold { font-weight: 700; }
.dr-d-none { display: none; }

/* ---- Back Button ---- */
.dr-back-btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 16px;
    color: var(--dr-primary);
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    border: none;
    background: none;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    min-height: 48px;
}

/* ---- Admin Scenario List ---- */
.dr-scenario-card {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 16px 20px;
    border: 1px solid var(--dr-border);
    border-radius: var(--dr-radius);
    margin-bottom: 12px;
    background: white;
    cursor: pointer;
    transition: all 0.15s;
    min-height: 60px;
    -webkit-tap-highlight-color: transparent;
}

.dr-scenario-card:active { background: var(--dr-light-bg); }

.dr-scenario-name {
    font-size: 20px;
    font-weight: 600;
}

.dr-scenario-meta {
    font-size: 14px;
    color: var(--dr-gray);
    margin-top: 4px;
}

.dr-scenario-actions {
    display: flex;
    gap: 8px;
}

/* ---- Admin Station Editor ---- */
.dr-station-item {
    padding: 16px;
    border: 1px solid var(--dr-border);
    border-radius: var(--dr-radius);
    margin-bottom: 12px;
    background: var(--dr-light-bg);
}

.dr-station-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 12px;
}

.dr-dest-item {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 12px;
    border: 1px solid var(--dr-border);
    border-radius: 8px;
    margin-bottom: 8px;
    background: white;
    flex-wrap: wrap;
}

.dr-dest-item .dr-input,
.dr-dest-item .dr-select {
    flex: 1;
    min-width: 120px;
}

/* ---- Search Picker ---- */
.dr-search-box {
    position: relative;
}

.dr-search-results {
    position: absolute;
    top: 100%;
    left: 0;
    right: 0;
    background: white;
    border: 1px solid var(--dr-border);
    border-radius: 0 0 8px 8px;
    box-shadow: var(--dr-shadow-lg);
    max-height: 50vh;
    overflow-y: auto;
    z-index: 1050;
}

.dr-search-result-item {
    padding: 12px 16px;
    cursor: pointer;
    border-bottom: 1px solid var(--dr-border);
    min-height: 48px;
    display: flex;
    align-items: center;
    -webkit-tap-highlight-color: transparent;
}

.dr-search-result-item:active { background: var(--dr-primary-light); }

.dr-search-result-item:last-child { border-bottom: none; }

.dr-search-result-name {
    font-weight: 600;
    font-size: 16px;
}

.dr-search-result-meta {
    font-size: 13px;
    color: var(--dr-gray);
}

.dr-selected-org {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 10px 16px;
    background: var(--dr-primary-light);
    border-radius: 8px;
    font-weight: 600;
}

.dr-selected-org .dr-remove-btn {
    cursor: pointer;
    color: var(--dr-danger);
    font-size: 18px;
    padding: 4px;
    min-width: 32px;
    min-height: 32px;
    display: flex;
    align-items: center;
    justify-content: center;
}

/* ---- Station Mode: Picker ---- */
.dr-picker-list {
    display: flex;
    flex-direction: column;
    gap: 12px;
}

.dr-picker-btn {
    display: flex;
    align-items: center;
    padding: 20px 24px;
    border: 2px solid var(--dr-border);
    border-radius: var(--dr-radius);
    background: white;
    cursor: pointer;
    font-size: 22px;
    font-weight: 600;
    min-height: 70px;
    transition: all 0.15s;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
}

.dr-picker-btn:active { background: var(--dr-light-bg); transform: scale(0.98); }

.dr-picker-btn i {
    font-size: 28px;
    margin-right: 16px;
    color: var(--dr-primary);
}

/* ---- Station Mode: Main Registration View ---- */
.dr-station-header-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 16px;
    background: var(--dr-dark);
    color: white;
    border-radius: var(--dr-radius);
    margin-bottom: 12px;
    flex-wrap: wrap;
    gap: 8px;
}

.dr-station-header-bar .dr-header-title {
    font-size: 20px;
    font-weight: 700;
}

.dr-station-header-bar .dr-header-subtitle {
    font-size: 14px;
    opacity: 0.8;
}

.dr-station-header-bar .dr-header-actions {
    display: flex;
    gap: 8px;
}

.dr-search-bar {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 12px;
}

.dr-search-wrap {
    flex: 1;
    position: relative;
}
.dr-search-wrap .dr-input { width: 100%; padding-right: 32px; }
.dr-search-clear {
    position: absolute;
    right: 8px;
    top: 50%;
    transform: translateY(-50%);
    border: none;
    background: none;
    color: #999;
    font-size: 18px;
    cursor: pointer;
    padding: 0 4px;
    line-height: 1;
    display: none;
}
.dr-search-clear:hover { color: #333; }

.dr-main-layout {
    display: flex;
    gap: 16px;
}

.dr-people-panel {
    flex: 1;
    min-width: 0;
}

.dr-dest-panel {
    width: 380px;
    flex-shrink: 0;
}

/* Person cards */
.dr-person-card {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 16px;
    border: 2px solid var(--dr-border);
    border-radius: 10px;
    margin-bottom: 8px;
    background: white;
    cursor: pointer;
    transition: all 0.15s;
    min-height: 56px;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
}

.dr-person-card:active { transform: scale(0.98); }
.dr-person-card.dr-selected {
    border-color: var(--dr-primary);
    background: var(--dr-primary-light);
    box-shadow: 0 0 0 3px rgba(37,99,235,0.2);
}
.dr-person-card.dr-assigned {
    opacity: 1;
    background: #f0fdf4;
}

.dr-person-name {
    font-size: 18px;
    font-weight: 600;
}

.dr-person-meta {
    font-size: 14px;
    color: var(--dr-gray);
}

.dr-person-badge {
    display: inline-flex;
    align-items: center;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
}

.dr-badge-assigned {
    background: var(--dr-secondary-light);
    color: var(--dr-secondary);
}

/* People list tabs */
.dr-people-tabs {
    display: flex;
    border-bottom: 2px solid #e0e0e0;
    margin: 0;
    padding: 0;
}
.dr-people-tab {
    flex: 1;
    padding: 10px 8px;
    text-align: center;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    border: none;
    background: none;
    color: var(--dr-gray);
    border-bottom: 3px solid transparent;
    margin-bottom: -2px;
    transition: all 0.15s;
}
.dr-people-tab:hover { color: var(--dr-dark); background: #f8f9fa; }
.dr-people-tab.dr-tab-active {
    color: var(--dr-primary);
    border-bottom-color: var(--dr-primary);
}
.dr-people-tab .dr-tab-count {
    display: inline-block;
    min-width: 22px;
    padding: 1px 6px;
    border-radius: 10px;
    font-size: 12px;
    font-weight: 700;
    margin-left: 4px;
    background: #e0e0e0;
    color: #555;
}
.dr-people-tab.dr-tab-active .dr-tab-count {
    background: var(--dr-primary);
    color: #fff;
}

/* Station switcher above destinations */
.dr-dest-station-switcher {
    padding: 10px 14px;
    margin-bottom: 12px;
    background: var(--dr-light-bg);
    border: 1px solid var(--dr-border);
    border-radius: var(--dr-radius);
    display: flex;
    align-items: center;
    gap: 6px;
    position: relative;
}
.dr-dest-station-label {
    font-size: 13px;
    color: var(--dr-gray);
    font-weight: 500;
}
.dr-dest-station-switcher .dr-station-switch-trigger {
    font-size: 16px;
    font-weight: 700;
    color: var(--dr-primary);
}

/* Destination cards */
.dr-dest-card {
    border: 2px solid var(--dr-dest-color, var(--dr-border));
    border-radius: var(--dr-radius);
    margin-bottom: 12px;
    overflow: hidden;
    background: white;
    cursor: pointer;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
    transition: all 0.15s;
}

.dr-dest-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 16px;
    background: var(--dr-dest-bg, var(--dr-light-bg));
    border-bottom: 1px solid var(--dr-dest-color, var(--dr-border));
}

.dr-dest-card-name {
    font-size: 18px;
    font-weight: 700;
    color: var(--dr-dest-color, var(--dr-dark));
}

.dr-dest-card-count {
    font-size: 22px;
    font-weight: 800;
    color: var(--dr-dest-color, var(--dr-dark));
}

.dr-dest-card-body { padding: 8px 16px 10px; }
.dr-dest-card-detail { padding: 0 16px 10px; display: none; }
.dr-dest-card.dr-dest-expanded .dr-dest-card-detail { display: block; }

/* Capacity bar */
.dr-cap-bar-wrap {
    height: 10px;
    background: #e5e7eb;
    border-radius: 5px;
    overflow: hidden;
    margin-bottom: 8px;
}

.dr-cap-bar {
    height: 100%;
    border-radius: 5px;
    transition: width 0.4s ease, background-color 0.3s;
}

.dr-cap-bar.dr-cap-ok { background: var(--dr-secondary); }
.dr-cap-bar.dr-cap-warn { background: var(--dr-warning); }
.dr-cap-bar.dr-cap-full { background: var(--dr-danger); }

.dr-dest-names {
    margin-top: 6px;
    font-size: 12px;
    color: #555;
    max-height: 120px;
    overflow-y: auto;
}
.dr-dest-name-item {
    padding: 1px 0;
    border-bottom: 1px solid rgba(0,0,0,0.05);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* Clickable destination cards */
.dr-dest-card:active { transform: scale(0.98); }
.dr-dest-card.dr-dest-disabled { opacity: 0.5; cursor: not-allowed; }
.dr-dest-card.dr-dest-disabled:active { transform: none; }

/* Selected destination card */
.dr-dest-card.dr-dest-selected {
    border-width: 3px;
    border-color: var(--dr-dest-color, var(--dr-primary));
    box-shadow: 0 0 0 3px rgba(59,130,246,0.25), 0 4px 12px rgba(0,0,0,0.15);
    transform: scale(1.02);
}

/* Floating assign action bar */
.dr-assign-action-bar {
    position: fixed;
    bottom: 0;
    left: 0;
    right: 0;
    background: white;
    border-top: 2px solid var(--dr-primary);
    box-shadow: 0 -4px 20px rgba(0,0,0,0.15);
    padding: 12px 20px;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 16px;
    z-index: 1000;
    animation: drSlideUp 0.25s ease;
}
@keyframes drSlideUp {
    from { transform: translateY(100%); }
    to { transform: translateY(0); }
}
.dr-assign-bar-info {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 16px;
    font-weight: 600;
    color: var(--dr-dark);
}
.dr-assign-bar-arrow {
    color: var(--dr-muted);
    font-size: 18px;
}
.dr-assign-bar-btn {
    padding: 14px 40px;
    border: none;
    border-radius: 10px;
    font-size: 20px;
    font-weight: 700;
    color: white;
    background: var(--dr-primary);
    cursor: pointer;
    min-height: 54px;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
    transition: all 0.15s;
}
.dr-assign-bar-btn:active { transform: scale(0.97); }

/* ---- Toast ---- */
.dr-toast-container {
    position: fixed;
    bottom: 20px;
    left: 50%;
    transform: translateX(-50%);
    z-index: 9999;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 8px;
    pointer-events: none;
}

.dr-toast {
    padding: 14px 28px;
    border-radius: 10px;
    font-size: 16px;
    font-weight: 600;
    color: white;
    opacity: 0;
    transform: translateY(20px);
    transition: all 0.3s;
    pointer-events: auto;
    max-width: 90vw;
    text-align: center;
}

.dr-toast.dr-show { opacity: 1; transform: translateY(0); }
.dr-toast-success { background: var(--dr-secondary); }
.dr-toast-warning { background: var(--dr-warning); }
.dr-toast-danger { background: var(--dr-danger); }
.dr-toast-info { background: var(--dr-primary); }

/* ---- Connection Banner ---- */
.dr-conn-banner {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    padding: 8px;
    background: var(--dr-warning);
    color: white;
    text-align: center;
    font-weight: 600;
    font-size: 14px;
    z-index: 9998;
    display: none;
}

/* ---- Modal ---- */
.dr-modal-overlay {
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0,0,0,0.5);
    z-index: 5000;
    display: none;
    align-items: center;
    justify-content: center;
    padding: 20px;
}

.dr-modal-overlay.dr-show { display: flex; }

.dr-modal {
    background: white;
    border-radius: var(--dr-radius);
    width: 100%;
    max-width: 700px;
    max-height: 85vh;
    overflow-y: auto;
    box-shadow: var(--dr-shadow-lg);
}

.dr-modal-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 16px 20px;
    border-bottom: 1px solid var(--dr-border);
    font-size: 20px;
    font-weight: 700;
}

.dr-modal-close {
    font-size: 24px;
    cursor: pointer;
    padding: 8px;
    min-width: 40px;
    min-height: 40px;
    display: flex;
    align-items: center;
    justify-content: center;
    border: none;
    background: none;
    -webkit-tap-highlight-color: transparent;
}

.dr-modal-body { padding: 20px; }
.dr-modal-footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    padding: 16px 20px;
    border-top: 1px solid var(--dr-border);
}

/* ---- Status Indicator ---- */
.dr-status-item {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 13px;
    white-space: nowrap;
}

.dr-status-dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    display: inline-block;
}

.dr-status-dot.dr-green { background: var(--dr-secondary); }
.dr-status-dot.dr-yellow { background: var(--dr-warning); }

/* ---- Empty State ---- */
.dr-empty {
    text-align: center;
    padding: 40px 20px;
    color: var(--dr-gray);
}

.dr-empty i {
    font-size: 48px;
    margin-bottom: 12px;
    display: block;
}

/* ---- Loading Spinner ---- */
.dr-loading {
    display: none;
    text-align: center;
    padding: 20px;
}

.dr-spinner {
    width: 40px;
    height: 40px;
    border: 4px solid var(--dr-border);
    border-top-color: var(--dr-primary);
    border-radius: 50%;
    animation: dr-spin 0.8s linear infinite;
    margin: 0 auto 12px;
}

@keyframes dr-spin { to { transform: rotate(360deg); } }

/* ---- Responsive ---- */
@media (max-width: 900px) {
    .dr-main-layout {
        flex-direction: column;
    }
    .dr-dest-panel {
        width: 100%;
    }
}

@media (max-width: 600px) {
    .dr-landing-btn {
        width: 100%;
        height: 150px;
    }
    .dr-container { padding: 8px; }
    .dr-person-name { font-size: 16px; }
    .dr-dest-card-name { font-size: 16px; }
}

/* ---- Person Info Bar ---- */
.dr-info-bar {
    background: white;
    border: 2px solid var(--dr-primary);
    border-radius: var(--dr-radius);
    box-shadow: var(--dr-shadow-lg);
    margin-bottom: 12px;
    overflow: hidden;
    animation: dr-slideDown 0.2s ease-out;
}

@keyframes dr-slideDown {
    from { opacity: 0; transform: translateY(-10px); }
    to { opacity: 1; transform: translateY(0); }
}

.dr-info-bar-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 16px;
    background: var(--dr-primary);
    color: white;
    font-size: 17px;
    font-weight: 700;
}

.dr-info-bar-dismiss {
    background: none;
    border: none;
    color: white;
    font-size: 22px;
    cursor: pointer;
    padding: 4px 8px;
    min-width: 40px;
    min-height: 40px;
    display: flex;
    align-items: center;
    justify-content: center;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
}

.dr-info-bar-body {
    padding: 10px 16px;
    display: flex;
    flex-wrap: wrap;
    gap: 8px 20px;
}

.dr-info-section {
    display: flex;
    align-items: baseline;
    gap: 6px;
    font-size: 15px;
    line-height: 1.4;
}

.dr-info-section-icon {
    flex-shrink: 0;
    font-size: 15px;
}

.dr-info-label {
    font-weight: 700;
    color: var(--dr-gray);
    white-space: nowrap;
}

.dr-info-value {
    color: var(--dr-dark);
}

.dr-info-emergency {
    border-left: 4px solid var(--dr-danger);
    padding-left: 10px;
}

.dr-info-clear {
    border-left: 4px solid var(--dr-secondary);
    padding-left: 10px;
}

.dr-info-medical-alert {
    background: var(--dr-danger-light);
    border-radius: 6px;
    padding: 6px 10px;
    color: var(--dr-danger);
    font-weight: 600;
    font-size: 15px;
}

.dr-info-medical-clear {
    color: var(--dr-secondary);
    font-size: 14px;
}

.dr-info-answers-toggle {
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
    font-size: 15px;
    font-weight: 600;
    color: var(--dr-primary);
    padding: 6px 0;
    -webkit-tap-highlight-color: transparent;
    touch-action: manipulation;
    user-select: none;
    border: none;
    background: none;
    min-height: 40px;
}

.dr-info-answers-content {
    display: none;
    padding: 6px 0;
}

.dr-info-answers-content.dr-show {
    display: block;
}

.dr-info-answer-row {
    display: flex;
    gap: 8px;
    padding: 4px 0;
    font-size: 14px;
    border-bottom: 1px solid var(--dr-border);
}

.dr-info-answer-row:last-child {
    border-bottom: none;
}

.dr-info-answer-q {
    font-weight: 600;
    color: var(--dr-gray);
    min-width: 120px;
    flex-shrink: 0;
}

.dr-info-answer-a {
    color: var(--dr-dark);
}

.dr-info-family {
    font-size: 14px;
    color: var(--dr-dark);
}

.dr-info-family-sep {
    color: var(--dr-gray);
    margin: 0 2px;
}

.dr-info-bar-footer {
    padding: 6px 16px 10px;
    border-top: 1px solid var(--dr-border);
}

@media (max-width: 600px) {
    .dr-info-bar-body {
        flex-direction: column;
        gap: 6px;
    }
    .dr-info-answer-row {
        flex-direction: column;
        gap: 2px;
    }
    .dr-info-answer-q {
        min-width: unset;
    }
}

/* ---- Fullscreen Mode ---- */
.dr-fullscreen.dr-container {
    max-width: 100%;
    padding: 0;
}
.dr-fullscreen .dr-station-header-bar {
    border-radius: 0;
}

/* ---- Quick Station Switcher ---- */
.dr-station-switch-trigger {
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    border-radius: 6px;
    padding: 2px 8px 2px 0;
    transition: background 0.15s;
}
.dr-station-switch-trigger:hover {
    background: rgba(255,255,255,0.15);
}
.dr-station-switch-trigger .bi-chevron-down {
    font-size: 0.7em;
    opacity: 0.7;
    transition: transform 0.2s;
}
.dr-station-switch-trigger.open .bi-chevron-down {
    transform: rotate(180deg);
}
.dr-station-switcher-wrap {
    position: relative;
    display: inline-block;
}
.dr-station-switcher {
    position: absolute;
    top: calc(100% + 6px);
    left: 0;
    min-width: 220px;
    background: #1e293b;
    border: 1px solid rgba(255,255,255,0.15);
    border-radius: 10px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.4);
    z-index: 9000;
    overflow: hidden;
    animation: drSwitcherIn 0.15s ease-out;
}
@keyframes drSwitcherIn {
    from { opacity: 0; transform: translateY(-6px); }
    to   { opacity: 1; transform: translateY(0); }
}
.dr-station-switcher-title {
    padding: 10px 14px 6px;
    font-size: 0.7em;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: rgba(255,255,255,0.45);
    font-weight: 600;
}
.dr-station-switcher-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 14px;
    color: rgba(255,255,255,0.85);
    font-size: 0.95em;
    cursor: pointer;
    transition: background 0.12s;
    border: none;
    background: none;
    width: 100%;
    text-align: left;
    font-family: inherit;
}
.dr-station-switcher-item:hover {
    background: rgba(255,255,255,0.1);
}
.dr-station-switcher-item.active {
    background: rgba(37,99,235,0.3);
    color: #fff;
    font-weight: 600;
}
.dr-station-switcher-item .bi-check-lg {
    color: #60a5fa;
    font-size: 1.1em;
}
.dr-station-switcher-backdrop {
    position: fixed;
    inset: 0;
    z-index: 8999;
}
/* Pinned Questions - Admin Editor */
.dr-pin-questions-section {
    border: 1px solid var(--dr-border);
    border-radius: 8px;
    padding: 12px;
    background: var(--dr-bg);
}
.dr-pin-question-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 0;
    border-bottom: 1px solid var(--dr-border);
}
.dr-pin-question-row:last-child {
    border-bottom: none;
}
.dr-input-sm {
    padding: 4px 8px;
    font-size: 13px;
    height: 28px;
}
.dr-dest-subgroups {
    padding: 4px 0 8px;
}
.dr-dest-subgroups .dr-mb-4 {
    margin-bottom: 4px;
}
.dr-subgroup-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 5px 12px;
    margin: 3px 4px 3px 0;
    border: 1.5px solid var(--dr-primary);
    border-radius: 14px;
    background: #fff;
    color: var(--dr-primary);
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.15s, color 0.15s;
}
.dr-subgroup-btn:hover {
    background: var(--dr-primary);
    color: #fff;
}
.dr-subgroup-btn.dr-sg-active {
    background: var(--dr-primary);
    color: #fff;
    border-color: var(--dr-primary);
}
.dr-subgroup-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}
.dr-subgroup-btn i {
    font-size: 11px;
}
    </style>
</head>
<body>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>
// ====================================================================
// Day of Registration - Application JavaScript
// ====================================================================
(function() {
    var scriptPath = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');

    // ---- Auto-update constants ----
    var APP_VERSION = ''' + json.dumps(APP_VERSION) + ''';
    var DC_SCRIPT_ID = ''' + json.dumps(DC_SCRIPT_ID) + ''';
    var DC_API_BASE = ''' + json.dumps(DC_API_BASE) + ''';
    // SCRIPT_NAME is detected client-side from the URL the browser actually loaded —
    // bulletproof regardless of how the admin renamed the install.
    var SCRIPT_NAME = (function() {
        try {
            var m = (window.location.pathname || '').match(/\\/PyScript(?:Form)?\\/([^\\/?#]+)/);
            if (m && m[1]) return m[1];
        } catch(e) {}
        return ''' + json.dumps(get_script_name()) + ''';
    })();
    var APP_UPDATE_AVAILABLE = false;
    var APP_LATEST_VERSION = '';

    // ---- State ----
    var state = {
        mode: 'landing',
        scenarios: [],
        currentScenario: null,
        currentStation: null,
        volunteerName: '',
        registrants: [],
        destinations: [],
        assignedIds: {},
        selectedPerson: null,
        selectedDestIdx: null,
        personDetail: null,
        searchTerm: '',
        peopleTab: 'pending',
        pollTimer: null,
        pollFailCount: 0,
        fullscreen: false,
        destSort: (function() { try { return localStorage.getItem('drDestSort') || 'name'; } catch(e) { return 'name'; } })()
    };

    // ---- AJAX Helper ----
    function extractJson(text) {
        // Try direct parse first
        text = (text || '').trim();
        try { return JSON.parse(text); } catch(e) {}
        // TouchPoint may wrap JSON in HTML - try to extract it
        var start = text.indexOf('{');
        var end = text.lastIndexOf('}');
        if (start >= 0 && end > start) {
            try { return JSON.parse(text.substring(start, end + 1)); } catch(e2) {}
        }
        return null;
    }

    function ajax(action, data, callback) {
        data = data || {};
        data.action = action;
        // Always echo the script name so server handlers (e.g. apply_update) write to the
        // right Python content slot regardless of how the admin renamed the install.
        if (typeof SCRIPT_NAME !== 'undefined' && SCRIPT_NAME) {
            data.script_name = SCRIPT_NAME;
        }
        console.log('[DR] AJAX =>', action, data);
        $.ajax({
            url: scriptPath,
            type: 'POST',
            data: data,
            dataType: 'text',
            success: function(response) {
                console.log('[DR] Response for', action, ':', response.substring(0, 300));
                var parsed = extractJson(response);
                if (parsed) {
                    callback(parsed);
                } else {
                    console.error('[DR] Could not parse response for', action, ':', response.substring(0, 500));
                    callback({success: false, message: 'Could not parse server response'});
                }
            },
            error: function(xhr, status, error) {
                console.error('[DR] AJAX error for', action, ':', status, error);
                callback({success: false, message: 'Request failed: ' + error});
            }
        });
    }

    // ---- Toast ----
    function showToast(msg, toastType) {
        toastType = toastType || 'info';
        var toast = $('<div class="dr-toast dr-toast-' + toastType + '">' + msg + '</div>');
        $('#drToastContainer').append(toast);
        setTimeout(function() { toast.addClass('dr-show'); }, 10);
        setTimeout(function() {
            toast.removeClass('dr-show');
            setTimeout(function() { toast.remove(); }, 300);
        }, 3000);
    }

    // ---- Modal ----
    function showModal(id) { $('#' + id).addClass('dr-show'); }
    function hideModal(id) { $('#' + id).removeClass('dr-show'); }

    // ---- Navigation ----
    function navigate(mode) {
        state.mode = mode;
        render();
    }

    // ---- Render Router ----
    function render() {
        var $app = $('#drApp');
        switch(state.mode) {
            case 'landing': renderLanding($app); break;
            case 'admin': renderAdmin($app); break;
            case 'admin_edit': renderAdminEdit($app); break;
            case 'station_pick_scenario': renderStationPickScenario($app); break;
            case 'station_pick_station': renderStationPickStation($app); break;
            case 'station_pick_volunteer': renderStationPickVolunteer($app); break;
            case 'station_main': renderStationMain($app); break;
            default: renderLanding($app);
        }
    }

    // ================================================================
    // LANDING PAGE
    // ================================================================
    function renderLanding($app) {
        $app.html(
            '<div class="dr-landing">' +
                '<div>' +
                    '<div class="dr-landing-title"><i class="bi bi-clipboard-check"></i> Day of Registration</div>' +
                    '<div class="dr-landing-subtitle">Event registration and room assignment tool</div>' +
                '</div>' +
                '<div class="dr-landing-buttons">' +
                    '<div class="dr-landing-btn dr-admin-btn" onclick="DRApp.goAdmin()">' +
                        '<i class="bi bi-gear-fill"></i>' +
                        '<span>Admin Setup</span>' +
                    '</div>' +
                    '<div class="dr-landing-btn dr-station-btn" onclick="DRApp.goStation()">' +
                        '<i class="bi bi-tablet-landscape"></i>' +
                        '<span>Station Mode</span>' +
                    '</div>' +
                '</div>' +
            '</div>'
        );
    }

    // ================================================================
    // ADMIN - Scenario Manager
    // ================================================================
    function renderAdmin($app) {
        var h = '<div class="dr-back-btn" onclick="DRApp.goLanding()"><i class="bi bi-arrow-left"></i> Back</div>';
        h += '<div class="dr-flex dr-items-center dr-justify-between dr-mb-16">';
        h += '<h2 style="margin:0">Scenarios</h2>';
        h += '<button class="dr-btn dr-btn-primary" onclick="DRApp.createScenario()"><i class="bi bi-plus-lg"></i> New Scenario</button>';
        h += '</div>';
        h += '<div id="drScenarioList"><div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Loading...</div></div>';
        $app.html(h);
        loadScenarios();
    }

    function loadScenarios() {
        ajax('load_scenarios', {}, function(resp) {
            if (resp.success) {
                state.scenarios = resp.scenarios || [];
                renderScenarioList();
            } else {
                $('#drScenarioList').html('<div class="dr-empty"><i class="bi bi-exclamation-triangle"></i>Error loading scenarios</div>');
            }
        });
    }

    function renderScenarioList() {
        var $list = $('#drScenarioList');
        if (state.scenarios.length === 0) {
            $list.html('<div class="dr-empty"><i class="bi bi-folder-plus"></i><div>No scenarios yet</div><div class="dr-text-sm">Create one to get started</div></div>');
            return;
        }
        var h = '';
        for (var i = 0; i < state.scenarios.length; i++) {
            var s = state.scenarios[i];
            var stationCount = (s.stations || []).length;
            h += '<div class="dr-scenario-card">';
            h += '<div onclick="DRApp.editScenario(\\'' + s.id + '\\')">';
            h += '<div class="dr-scenario-name">' + escHtml(s.name || 'Unnamed') + '</div>';
            h += '<div class="dr-scenario-meta">' + stationCount + ' station' + (stationCount !== 1 ? 's' : '') + ' &middot; Updated ' + escHtml(s.updatedAt || '') + '</div>';
            h += '</div>';
            h += '<div class="dr-scenario-actions">';
            h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="event.stopPropagation(); DRApp.editScenario(\\'' + s.id + '\\')"><i class="bi bi-pencil"></i></button>';
            h += '<button class="dr-btn dr-btn-danger dr-btn-sm" onclick="event.stopPropagation(); DRApp.deleteScenario(\\'' + s.id + '\\', \\'' + escHtml(s.name || '') + '\\')"><i class="bi bi-trash"></i></button>';
            h += '</div></div>';
        }
        $list.html(h);
    }

    function createScenario() {
        state.currentScenario = {
            id: '',
            name: '',
            friendFinderProgramId: '',
            friendFinderDivisionId: '',
            stations: []
        };
        state.mode = 'admin_edit';
        render();
    }

    function editScenario(id) {
        for (var i = 0; i < state.scenarios.length; i++) {
            if (state.scenarios[i].id === id) {
                state.currentScenario = JSON.parse(JSON.stringify(state.scenarios[i]));
                state.mode = 'admin_edit';
                render();
                return;
            }
        }
    }

    function deleteScenario(id, name) {
        if (!confirm('Delete scenario "' + name + '"? This cannot be undone.')) return;
        ajax('delete_scenario', {scenario_id: id}, function(resp) {
            if (resp.success) {
                showToast('Scenario deleted', 'success');
                loadScenarios();
            } else {
                showToast('Error: ' + resp.message, 'danger');
            }
        });
    }

    // ================================================================
    // ADMIN - Scenario Editor
    // ================================================================
    function renderAdminEdit($app) {
        var sc = state.currentScenario;
        var h = '<div class="dr-back-btn" onclick="DRApp.goAdminList()"><i class="bi bi-arrow-left"></i> Back to Scenarios</div>';
        h += '<div class="dr-panel"><div class="dr-panel-header"><span><i class="bi bi-pencil-square"></i> ' + (sc.id ? 'Edit' : 'New') + ' Scenario</span>';
        h += '<button class="dr-btn dr-btn-primary dr-btn-sm" onclick="DRApp.saveScenario()"><i class="bi bi-check-lg"></i> Save</button></div>';
        h += '<div class="dr-panel-body">';
        h += '<div class="dr-mb-16"><label class="dr-label">Scenario Name</label>';
        h += '<input type="text" class="dr-input" id="drScenarioName" value="' + escAttr(sc.name || '') + '" placeholder="e.g., VBS 2025"></div>';

        // Friend Finder Filter
        h += '<div class="dr-mb-16" style="padding:12px;background:#f8f9fa;border-radius:8px;border:1px solid #e0e0e0">';
        h += '<label class="dr-label" style="margin-bottom:8px"><i class="bi bi-search-heart"></i> Find Friend Filter</label>';
        h += '<div class="dr-text-sm dr-text-muted" style="margin-bottom:8px">Optionally restrict Find Friend results to people in involvements within a specific program/division.</div>';
        h += '<div class="dr-flex dr-gap-8" style="flex-wrap:wrap">';
        h += '<div style="flex:1;min-width:180px"><label class="dr-label dr-text-sm">Program</label>';
        h += '<select class="dr-input" id="drFFProgram" onchange="DRApp.ffProgramChanged()"><option value="">All Programs</option></select></div>';
        h += '<div style="flex:1;min-width:180px"><label class="dr-label dr-text-sm">Division</label>';
        h += '<select class="dr-input" id="drFFDivision"><option value="">All Divisions</option></select></div>';
        h += '</div></div>';

        // Test Mode + cleanup tools
        var isTestMode = !!sc.isTestMode;
        h += '<div class="dr-mb-16" style="padding:12px;background:' + (isTestMode ? '#fff7ed' : '#f8f9fa') + ';border-radius:8px;border:1px solid ' + (isTestMode ? '#fb923c' : '#e0e0e0') + '">';
        h += '<label class="dr-label" style="margin-bottom:6px;cursor:pointer">';
        h += '<input type="checkbox" id="drIsTestMode"' + (isTestMode ? ' checked' : '') + ' onchange="DRApp.toggleTestMode(this.checked)"> ';
        h += '<i class="bi bi-flask"></i> <strong>Test Mode</strong>';
        h += '</label>';
        h += '<div class="dr-text-sm dr-text-muted" style="margin-bottom:8px">When enabled, you can use the destructive <em>Reset Test Assignments</em> action to wipe everyone out of every destination room and clear the log. Disable before going live.</div>';
        h += '<div class="dr-flex dr-gap-8" style="flex-wrap:wrap">';
        h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.clearAssignmentLog()" title="Empty the audit log for this scenario. Does not touch any people\\'s involvement membership."><i class="bi bi-eraser"></i> Clear Assignment Log</button>';
        if (isTestMode) {
            h += '<button class="dr-btn dr-btn-danger dr-btn-sm" onclick="DRApp.resetTestAssignments()" title="Remove every assigned person from every destination room. Backs up first."><i class="bi bi-exclamation-triangle"></i> Reset Test Assignments</button>';
        }
        h += '</div></div>';

        // Stations
        h += '<div class="dr-flex dr-items-center dr-justify-between dr-mb-12">';
        h += '<h3 style="margin:0">Stations</h3>';
        h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.addStation()"><i class="bi bi-plus-lg"></i> Add Station</button>';
        h += '</div>';
        h += '<div id="drStationList">';
        h += renderStationsList(sc.stations || []);
        h += '</div>';

        h += '</div></div>';

        // Station editor modal
        h += renderStationEditorModal();

        $app.html(h);
        loadFFFilters();
    }

    // Friend Finder filter state
    var ffPrograms = [];
    var ffDivisions = [];

    function loadFFFilters() {
        ajax('get_filters', {}, function(resp) {
            if (resp.success) {
                ffPrograms = resp.programs || [];
                ffDivisions = resp.divisions || [];
                var sc = state.currentScenario;
                var savedProg = sc.friendFinderProgramId || '';
                var savedDiv = sc.friendFinderDivisionId || '';

                var $prog = $('#drFFProgram');
                $prog.empty().append('<option value="">All Programs</option>');
                for (var i = 0; i < ffPrograms.length; i++) {
                    var sel = (String(ffPrograms[i].id) === String(savedProg)) ? ' selected' : '';
                    $prog.append('<option value="' + ffPrograms[i].id + '"' + sel + '>' + escHtml(ffPrograms[i].name) + '</option>');
                }

                populateFFDivisions(savedProg, savedDiv);
            }
        });
    }

    function populateFFDivisions(progId, selectedDivId) {
        var $div = $('#drFFDivision');
        $div.empty().append('<option value="">All Divisions</option>');
        for (var i = 0; i < ffDivisions.length; i++) {
            var d = ffDivisions[i];
            if (!progId || String(d.progId) === String(progId)) {
                var sel = (String(d.id) === String(selectedDivId)) ? ' selected' : '';
                $div.append('<option value="' + d.id + '"' + sel + '>' + escHtml(d.name) + '</option>');
            }
        }
    }

    function ffProgramChanged() {
        var progId = $('#drFFProgram').val();
        populateFFDivisions(progId, '');
    }

    function renderStationsList(stations) {
        if (stations.length === 0) {
            return '<div class="dr-empty"><i class="bi bi-layers"></i><div>No stations</div></div>';
        }
        var h = '';
        for (var i = 0; i < stations.length; i++) {
            var st = stations[i];
            var destCount = (st.destinations || []).length;
            h += '<div class="dr-station-item">';
            h += '<div class="dr-station-header">';
            h += '<div><div class="dr-fw-bold dr-text-lg">' + escHtml(st.name || 'Unnamed Station') + '</div>';
            h += '<div class="dr-text-sm dr-text-muted">Source: ' + escHtml(st.sourceOrgName || 'Not set') + ' &middot; ' + destCount + ' destination' + (destCount !== 1 ? 's' : '') + '</div></div>';
            h += '<div class="dr-flex dr-gap-8">';
            h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.editStation(' + i + ')"><i class="bi bi-pencil"></i> Edit</button>';
            h += '<button class="dr-btn dr-btn-danger dr-btn-sm" onclick="DRApp.removeStation(' + i + ')"><i class="bi bi-trash"></i></button>';
            h += '</div></div></div>';
        }
        return h;
    }

    function renderStationEditorModal() {
        var h = '<div class="dr-modal-overlay" id="drStationModal">';
        h += '<div class="dr-modal">';
        h += '<div class="dr-modal-header"><span id="drStationModalTitle">Edit Station</span>';
        h += '<button class="dr-modal-close" onclick="DRApp.hideModal(\\'drStationModal\\')">&times;</button></div>';
        h += '<div class="dr-modal-body" id="drStationModalBody"></div>';
        h += '<div class="dr-modal-footer">';
        h += '<button class="dr-btn dr-btn-outline" onclick="DRApp.hideModal(\\'drStationModal\\')">Cancel</button>';
        h += '<button class="dr-btn dr-btn-primary" onclick="DRApp.saveStation()"><i class="bi bi-check-lg"></i> Save Station</button>';
        h += '</div></div></div>';
        return h;
    }

    var editingStationIdx = -1;
    var editingStation = null;

    function addStation() {
        editingStationIdx = -1;
        loadedRegQuestions = [];
        editingStation = {
            id: 'station_' + Date.now() + Math.floor(Math.random() * 1000),
            name: '',
            sourceType: 'involvement',
            sourceOrgId: 0,
            sourceOrgName: '',
            pinnedQuestions: [],
            destinations: []
        };
        renderStationEditor();
        showModal('drStationModal');
    }

    function editStation(idx) {
        editingStationIdx = idx;
        loadedRegQuestions = [];
        editingStation = JSON.parse(JSON.stringify(state.currentScenario.stations[idx]));
        renderStationEditor();
        showModal('drStationModal');
    }

    function removeStation(idx) {
        if (!confirm('Remove this station?')) return;
        state.currentScenario.stations.splice(idx, 1);
        $('#drStationList').html(renderStationsList(state.currentScenario.stations));
    }

    function renderStationEditor() {
        var st = editingStation;
        var h = '<div class="dr-mb-16"><label class="dr-label">Station Name</label>';
        h += '<input type="text" class="dr-input" id="drStationName" value="' + escAttr(st.name || '') + '" placeholder="e.g., Kindergarten" oninput="DRApp.updateStationName(this.value)"></div>';

        // Source involvement
        h += '<div class="dr-mb-16"><label class="dr-label">Source Involvement (Pre-Registration List)</label>';
        if (st.sourceOrgId) {
            h += '<div class="dr-selected-org"><span>' + escHtml(st.sourceOrgName) + ' (#' + st.sourceOrgId + ')</span>';
            h += '<span class="dr-remove-btn" onclick="DRApp.clearStationSource()"><i class="bi bi-x-lg"></i></span></div>';
        } else {
            h += '<div class="dr-search-box"><input type="text" class="dr-input" id="drSourceSearch" placeholder="Search involvements..." oninput="DRApp.searchSource(this.value)">';
            h += '<div class="dr-search-results" id="drSourceResults" style="display:none"></div></div>';
        }
        h += '</div>';

        // Pinned Questions (only when source org is set)
        if (st.sourceOrgId) {
            h += '<div class="dr-mb-16 dr-pin-questions-section">';
            h += '<div class="dr-flex dr-items-center dr-justify-between dr-mb-8">';
            h += '<label class="dr-label" style="margin:0">Pinned Registration Questions</label>';
            h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.loadRegQuestions()"><i class="bi bi-arrow-clockwise"></i> Load Questions</button>';
            h += '</div>';
            h += '<div class="dr-text-sm dr-text-muted dr-mb-8">Pinned questions show immediately in the info bar instead of the collapsible section.</div>';
            h += '<div id="drPinQuestionsList">';
            h += renderPinQuestionsList();
            h += '</div></div>';
        }

        // Destinations
        h += '<div class="dr-flex dr-items-center dr-justify-between dr-mb-12">';
        h += '<label class="dr-label" style="margin:0">Destinations</label>';
        h += '<div class="dr-flex dr-gap-8">';
        h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.addDestBulk()"><i class="bi bi-list-check"></i> Bulk Add</button>';
        h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.addDest()"><i class="bi bi-plus-lg"></i> Add</button>';
        h += '</div>';
        h += '</div>';
        h += '<div id="drDestList">';
        h += renderDestList(st.destinations || []);
        h += '</div>';

        $('#drStationModalBody').html(h);
        $('#drStationModalTitle').text(editingStationIdx >= 0 ? 'Edit Station' : 'New Station');
    }

    // -- Pinned Questions helpers --
    var loadedRegQuestions = [];

    function renderPinQuestionsList() {
        var pinned = editingStation.pinnedQuestions || [];
        if (loadedRegQuestions.length === 0 && pinned.length === 0) {
            return '<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px">Click "Load Questions" to see available registration questions.</div>';
        }
        // Merge loaded questions with any already-pinned ones not in the loaded list
        var allLabels = [];
        var labelSet = {};
        for (var i = 0; i < loadedRegQuestions.length; i++) {
            var lbl = loadedRegQuestions[i].label;
            if (!labelSet[lbl]) { allLabels.push(lbl); labelSet[lbl] = true; }
        }
        for (var j = 0; j < pinned.length; j++) {
            if (!labelSet[pinned[j].question]) { allLabels.push(pinned[j].question); labelSet[pinned[j].question] = true; }
        }
        if (allLabels.length === 0) {
            return '<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px">No registration questions found for this involvement.</div>';
        }
        // Build pinned lookup
        var pinnedMap = {};
        for (var k = 0; k < pinned.length; k++) {
            pinnedMap[pinned[k].question] = pinned[k].shortLabel || '';
        }
        var h = '';
        for (var m = 0; m < allLabels.length; m++) {
            var label = allLabels[m];
            var isPinned = pinnedMap.hasOwnProperty(label);
            var shortLabel = isPinned ? pinnedMap[label] : '';
            h += '<div class="dr-pin-question-row">';
            h += '<label class="dr-flex dr-items-center dr-gap-8" style="cursor:pointer;flex:1;min-width:0">';
            h += '<input type="checkbox"' + (isPinned ? ' checked' : '') + ' data-label="' + escAttr(label) + '" onchange="DRApp.togglePinQuestion(this.dataset.label)">';
            h += '<span class="dr-text-sm" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escHtml(label) + '</span>';
            h += '</label>';
            if (isPinned) {
                h += '<input type="text" class="dr-input dr-input-sm" style="width:120px;flex-shrink:0" placeholder="Short label" value="' + escAttr(shortLabel) + '" data-label="' + escAttr(label) + '" onchange="DRApp.updatePinLabel(this.dataset.label, this.value)">';
            }
            h += '</div>';
        }
        return h;
    }

    function loadRegQuestions() {
        if (!editingStation || !editingStation.sourceOrgId) {
            showToast('Set a source involvement first', 'warning');
            return;
        }
        $('#drPinQuestionsList').html('<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px"><div class="dr-spinner" style="width:16px;height:16px;border-width:2px;display:inline-block;vertical-align:middle"></div> Loading...</div>');
        ajax('load_reg_questions', {
            org_id: editingStation.sourceOrgId
        }, function(resp) {
            if (resp.success) {
                loadedRegQuestions = resp.questions || [];
                $('#drPinQuestionsList').html(renderPinQuestionsList());
            } else {
                $('#drPinQuestionsList').html('<div class="dr-text-sm" style="color:var(--dr-danger);padding:8px">' + escHtml(resp.message || 'Failed to load') + '</div>');
            }
        });
    }

    function togglePinQuestion(label) {
        if (!editingStation) return;
        if (!editingStation.pinnedQuestions) editingStation.pinnedQuestions = [];
        var idx = -1;
        for (var i = 0; i < editingStation.pinnedQuestions.length; i++) {
            if (editingStation.pinnedQuestions[i].question === label) { idx = i; break; }
        }
        if (idx >= 0) {
            editingStation.pinnedQuestions.splice(idx, 1);
        } else {
            editingStation.pinnedQuestions.push({ question: label, shortLabel: '' });
        }
        $('#drPinQuestionsList').html(renderPinQuestionsList());
    }

    function updatePinLabel(label, shortLabel) {
        if (!editingStation || !editingStation.pinnedQuestions) return;
        for (var i = 0; i < editingStation.pinnedQuestions.length; i++) {
            if (editingStation.pinnedQuestions[i].question === label) {
                editingStation.pinnedQuestions[i].shortLabel = shortLabel;
                break;
            }
        }
    }

    function renderDestList(dests) {
        if (dests.length === 0) {
            return '<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:12px">No destinations added</div>';
        }
        var h = '';
        for (var i = 0; i < dests.length; i++) {
            var d = dests[i];
            h += '<div class="dr-dest-item">';
            h += '<div style="flex:1;min-width:200px">';
            h += '<div class="dr-fw-bold">' + escHtml(d.displayName || d.orgName || 'Unnamed') + '</div>';
            var sgInfo = '';
            if (d.subgroups && d.subgroups.length > 0) {
                sgInfo = ' &rarr; ' + d.subgroups.length + ' subgroup' + (d.subgroups.length > 1 ? 's' : '');
            } else if (d.subgroup) {
                sgInfo = ' &rarr; ' + escHtml(d.subgroup);
            }
            h += '<div class="dr-text-sm dr-text-muted">' + escHtml(d.orgName || 'No org') + sgInfo + '</div>';
            h += '</div>';
            h += '<div style="white-space:nowrap">';
            h += '<span class="dr-text-sm">Cap: ' + (d.capacity || 'None') + ' (' + (d.capType || 'soft') + ')</span>';
            h += '</div>';
            h += '<div class="dr-flex dr-gap-8">';
            h += '<button class="dr-btn dr-btn-outline dr-btn-sm" onclick="DRApp.editDest(' + i + ')"><i class="bi bi-pencil"></i></button>';
            h += '<button class="dr-btn dr-btn-danger dr-btn-sm" onclick="DRApp.removeDest(' + i + ')"><i class="bi bi-trash"></i></button>';
            h += '</div></div>';
        }
        return h;
    }

    var editingDestIdx = -1;

    function addDest() {
        editingDestIdx = -1;
        showDestEditor({
            id: 'dest_' + Date.now() + Math.floor(Math.random() * 1000),
            displayName: '',
            orgId: 0,
            orgName: '',
            subgroup: '',
            subgroups: [],
            capacity: 25,
            capType: 'soft'
        });
    }

    function editDest(idx) {
        editingDestIdx = idx;
        showDestEditor(JSON.parse(JSON.stringify(editingStation.destinations[idx])));
    }

    function removeDest(idx) {
        editingStation.destinations.splice(idx, 1);
        $('#drDestList').html(renderDestList(editingStation.destinations));
    }

    // ---- Bulk add destinations ----
    var bulkAddState = null;
    var bulkSearchTimer = null;
    var bulkLastResults = [];

    function addDestBulk() {
        bulkAddState = {
            defaultCapacity: 25,
            capType: 'soft',
            selected: {}  // orgId -> orgName
        };
        bulkLastResults = [];
        showDestBulkEditor();
    }

    function showDestBulkEditor() {
        var h = '<div class="dr-panel"><div class="dr-panel-header">Bulk Add Destinations</div>';
        h += '<div class="dr-panel-body">';
        h += '<div class="dr-text-sm dr-text-muted dr-mb-12">Search and check involvements to add. They\\'ll all be created with the default capacity below \\u2014 use the pencil icon afterward to fine-tune individual rooms or set subgroups.</div>';

        h += '<div class="dr-flex dr-gap-12 dr-mb-12">';
        h += '<div style="flex:1"><label class="dr-label">Default Capacity</label>';
        h += '<input type="number" class="dr-input" id="drBulkDefaultCap" value="' + bulkAddState.defaultCapacity + '" min="0" oninput="DRApp.updateBulkDefault()"></div>';
        h += '<div style="flex:1"><label class="dr-label">Cap Type</label>';
        h += '<select class="dr-select" id="drBulkCapType" onchange="DRApp.updateBulkDefault()">';
        h += '<option value="soft"' + (bulkAddState.capType === 'soft' ? ' selected' : '') + '>Soft (warn)</option>';
        h += '<option value="hard"' + (bulkAddState.capType === 'hard' ? ' selected' : '') + '>Hard (block)</option>';
        h += '</select></div></div>';

        h += '<div class="dr-mb-12"><label class="dr-label">Selected (<span id="drBulkSelCount">0</span>)</label>';
        h += '<div id="drBulkSelectedList" style="max-height:90px;overflow-y:auto;border:1px solid var(--dr-border);border-radius:6px;padding:6px;background:#fafafa"></div></div>';

        h += '<div class="dr-mb-12"><label class="dr-label">Search Involvements</label>';
        h += '<input type="text" class="dr-input dr-mb-8" id="drBulkSearch" placeholder="Type at least 2 characters..." oninput="DRApp.searchDestOrgBulk(this.value)">';
        h += '<div id="drBulkSearchResults" style="position:static;max-height:50vh;min-height:200px;overflow-y:auto;border:1px solid var(--dr-border);border-radius:6px;background:white"><div style="padding:12px;color:#999">Type to search...</div></div>';
        h += '</div>';

        h += '<div class="dr-flex dr-gap-8 dr-justify-between">';
        h += '<button class="dr-btn dr-btn-outline" onclick="DRApp.cancelBulkAdd()">Cancel</button>';
        h += '<button class="dr-btn dr-btn-primary" id="drBulkAddBtn" onclick="DRApp.commitBulkAdd()"><i class="bi bi-check-lg"></i> Add <span id="drBulkAddBtnCount">0</span> Selected</button>';
        h += '</div>';
        h += '</div></div>';

        $('#drDestList').html(h);
        renderBulkSelected();
    }

    function updateBulkDefault() {
        if (!bulkAddState) return;
        var v = parseInt($('#drBulkDefaultCap').val(), 10);
        bulkAddState.defaultCapacity = isNaN(v) ? 0 : v;
        bulkAddState.capType = $('#drBulkCapType').val() || 'soft';
    }

    function searchDestOrgBulk(term) {
        clearTimeout(bulkSearchTimer);
        if (term.length < 2) {
            bulkLastResults = [];
            $('#drBulkSearchResults').html('<div style="padding:12px;color:#999">Type to search...</div>');
            return;
        }
        bulkSearchTimer = setTimeout(function() {
            ajax('search_involvements', {search_term: term}, function(resp) {
                if (resp.success) {
                    bulkLastResults = resp.involvements || [];
                    renderBulkSearchResults();
                }
            });
        }, 300);
    }

    function renderBulkSearchResults() {
        var $el = $('#drBulkSearchResults');
        if (!bulkLastResults || bulkLastResults.length === 0) {
            $el.html('<div style="padding:12px;color:#999">No results</div>');
            return;
        }
        // Build set of orgIds already in destinations
        var existing = {};
        var existingDests = (editingStation && editingStation.destinations) ? editingStation.destinations : [];
        for (var i = 0; i < existingDests.length; i++) {
            if (existingDests[i].orgId) existing[existingDests[i].orgId] = true;
        }
        var h = '';
        for (var j = 0; j < bulkLastResults.length; j++) {
            var inv = bulkLastResults[j];
            var alreadyAdded = !!existing[inv.orgId];
            var picked = !!bulkAddState.selected[inv.orgId];
            var rowStyle = alreadyAdded ? 'opacity:0.5;cursor:not-allowed' : 'cursor:pointer';
            var bg = picked ? 'background:#e3f2fd' : '';
            h += '<div class="dr-search-result-item" style="padding:8px 12px;min-height:auto;' + rowStyle + ';' + bg + '" data-org-id="' + inv.orgId + '" data-org-name="' + escAttr(inv.orgName) + '" data-already="' + (alreadyAdded ? '1' : '0') + '" onclick="DRApp.toggleBulkPick(this)">';
            h += '<div class="dr-flex dr-items-center dr-gap-8" style="flex:1;min-width:0">';
            h += '<input type="checkbox"' + (picked ? ' checked' : '') + (alreadyAdded ? ' disabled' : '') + ' style="pointer-events:none;flex-shrink:0">';
            h += '<div style="flex:1;min-width:0">';
            h += '<div style="font-weight:600;font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + escHtml(inv.orgName) + (alreadyAdded ? ' <span class="dr-text-sm dr-text-muted">(already added)</span>' : '') + '</div>';
            h += '<div style="font-size:12px;color:var(--dr-gray);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">#' + inv.orgId + ' &middot; ' + escHtml(inv.programName) + ' &rarr; ' + escHtml(inv.divisionName) + ' &middot; ' + inv.memberCount + ' members</div>';
            h += '</div></div></div>';
        }
        $el.html(h);
    }

    function toggleBulkPick(el) {
        if (!bulkAddState) return;
        if (el.getAttribute('data-already') === '1') return;
        var orgId = parseInt(el.getAttribute('data-org-id'), 10);
        var orgName = el.getAttribute('data-org-name') || '';
        if (!orgId) return;
        if (bulkAddState.selected[orgId]) {
            delete bulkAddState.selected[orgId];
        } else {
            bulkAddState.selected[orgId] = orgName;
        }
        renderBulkSearchResults();
        renderBulkSelected();
    }

    function removeBulkPick(orgId) {
        if (!bulkAddState) return;
        delete bulkAddState.selected[orgId];
        renderBulkSelected();
        renderBulkSearchResults();
    }

    function renderBulkSelected() {
        var sel = bulkAddState ? bulkAddState.selected : {};
        var ids = [];
        for (var k in sel) { if (sel.hasOwnProperty(k)) ids.push(k); }
        $('#drBulkSelCount').text(ids.length);
        $('#drBulkAddBtnCount').text(ids.length);
        if (ids.length === 0) {
            $('#drBulkSelectedList').html('<div class="dr-text-sm dr-text-muted" style="padding:6px">None selected yet \\u2014 search below.</div>');
            return;
        }
        var h = '<div style="display:flex;flex-wrap:wrap;gap:4px">';
        for (var i = 0; i < ids.length; i++) {
            var oid = ids[i];
            var nm = sel[oid] || '';
            h += '<span title="' + escAttr(nm) + ' (#' + oid + ')" style="display:inline-flex;align-items:center;gap:4px;background:var(--dr-primary-light);border-radius:12px;padding:2px 8px;font-size:0.8em;max-width:240px">';
            h += '<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">#' + oid + ' ' + escHtml(nm) + '</span>';
            h += '<span data-org-id="' + oid + '" onclick="DRApp.removeBulkPick(parseInt(this.getAttribute(\\'data-org-id\\'),10))" style="cursor:pointer;color:var(--dr-danger);font-weight:bold;flex-shrink:0">\\u00d7</span>';
            h += '</span>';
        }
        h += '</div>';
        $('#drBulkSelectedList').html(h);
    }

    function cancelBulkAdd() {
        bulkAddState = null;
        bulkLastResults = [];
        $('#drDestList').html(renderDestList(editingStation.destinations));
    }

    function commitBulkAdd() {
        if (!bulkAddState) return;
        updateBulkDefault();
        var sel = bulkAddState.selected;
        var ids = [];
        for (var k in sel) { if (sel.hasOwnProperty(k)) ids.push(k); }
        if (ids.length === 0) {
            showToast('Select at least one involvement', 'warning');
            return;
        }
        if (!editingStation.destinations) editingStation.destinations = [];
        var added = 0;
        for (var i = 0; i < ids.length; i++) {
            var oid = parseInt(ids[i], 10);
            if (!oid) continue;
            // Skip if somehow already in destinations
            var dup = false;
            for (var j = 0; j < editingStation.destinations.length; j++) {
                if (editingStation.destinations[j].orgId === oid) { dup = true; break; }
            }
            if (dup) continue;
            editingStation.destinations.push({
                id: 'dest_' + Date.now() + '_' + Math.floor(Math.random() * 100000),
                displayName: '',
                orgId: oid,
                orgName: sel[ids[i]],
                subgroup: '',
                subgroups: [],
                capacity: bulkAddState.defaultCapacity,
                capType: bulkAddState.capType
            });
            added++;
        }
        bulkAddState = null;
        bulkLastResults = [];
        $('#drDestList').html(renderDestList(editingStation.destinations));
        showToast('Added ' + added + ' destination' + (added === 1 ? '' : 's'), 'success');
    }

    function showDestEditor(dest) {
        var h = '<div class="dr-panel"><div class="dr-panel-header">' + (editingDestIdx >= 0 ? 'Edit' : 'Add') + ' Destination</div>';
        h += '<div class="dr-panel-body">';
        h += '<div class="dr-mb-12"><label class="dr-label">Display Name <span class="dr-text-sm dr-text-muted">(optional - defaults to involvement name)</span></label>';
        h += '<input type="text" class="dr-input" id="drDestDisplayName" value="' + escAttr(dest.displayName || '') + '" placeholder="Leave blank to use involvement name"></div>';

        h += '<div class="dr-mb-12"><label class="dr-label">Destination Involvement</label>';
        if (dest.orgId) {
            h += '<div class="dr-selected-org" id="drDestOrgSelected"><span>' + escHtml(dest.orgName) + ' (#' + dest.orgId + ')</span>';
            h += '<span class="dr-remove-btn" onclick="DRApp.clearDestOrg()"><i class="bi bi-x-lg"></i></span></div>';
        } else {
            h += '<div class="dr-search-box"><input type="text" class="dr-input" id="drDestOrgSearch" placeholder="Search involvements..." oninput="DRApp.searchDestOrg(this.value)">';
            h += '<div class="dr-search-results" id="drDestOrgResults" style="display:none"></div></div>';
        }
        h += '<input type="hidden" id="drDestOrgId" value="' + (dest.orgId || 0) + '">';
        h += '<input type="hidden" id="drDestOrgName" value="' + escAttr(dest.orgName || '') + '">';
        h += '</div>';

        h += '<input type="hidden" id="drDestSubgroup" value="' + escAttr(dest.subgroup || '') + '">';

        // Subgroups section
        h += '<div class="dr-mb-12" id="drDestSubgroupsSection"' + (dest.orgId ? '' : ' style="display:none"') + '>';
        h += '<label class="dr-label">Subgroups to Show in Station Mode</label>';
        h += '<div class="dr-text-sm dr-text-muted dr-mb-8">Select which subgroups volunteers can assign people to from this destination.</div>';
        h += '<div id="drDestSubgroupsList">';
        if (dest.orgId) {
            h += '<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px"><div class="dr-spinner" style="width:16px;height:16px;border-width:2px;display:inline-block;vertical-align:middle"></div> Loading subgroups...</div>';
        } else {
            h += '<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px">Select an involvement first</div>';
        }
        h += '</div></div>';

        h += '<div class="dr-flex dr-gap-12 dr-mb-12">';
        h += '<div style="flex:1"><label class="dr-label">Capacity</label>';
        h += '<input type="number" class="dr-input" id="drDestCapacity" value="' + (dest.capacity || 25) + '" min="0"></div>';
        h += '<div style="flex:1"><label class="dr-label">Cap Type</label>';
        h += '<select class="dr-select" id="drDestCapType">';
        h += '<option value="soft"' + (dest.capType === 'soft' ? ' selected' : '') + '>Soft (warn)</option>';
        h += '<option value="hard"' + (dest.capType === 'hard' ? ' selected' : '') + '>Hard (block)</option>';
        h += '</select></div></div>';

        h += '<div class="dr-flex dr-gap-8 dr-justify-between">';
        h += '<button class="dr-btn dr-btn-outline" onclick="DRApp.cancelDestEdit()">Cancel</button>';
        h += '<button class="dr-btn dr-btn-primary" onclick="DRApp.saveDest()"><i class="bi bi-check-lg"></i> Save Destination</button>';
        h += '</div>';
        h += '</div></div>';

        // Store the dest data for saving
        window._editingDest = dest;
        $('#drDestList').html(h);

        // Auto-load subgroups if org already set
        if (dest.orgId) {
            loadDestSubgroups(dest.orgId);
        }
    }

    function loadDestSubgroups(orgId) {
        ajax('get_org_subgroups', { org_id: orgId }, function(resp) {
            if (resp.success) {
                window._availableSubgroups = resp.subgroups || [];
                renderDestSubgroupsList();
            } else {
                $('#drDestSubgroupsList').html('<div class="dr-text-sm" style="color:var(--dr-danger);padding:8px">Failed to load subgroups</div>');
            }
        });
    }

    function renderDestSubgroupsList() {
        var available = window._availableSubgroups || [];
        if (available.length === 0) {
            $('#drDestSubgroupsList').html('<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px">No subgroups found for this involvement.</div>');
            return;
        }
        var selected = (window._editingDest && window._editingDest.subgroups) ? window._editingDest.subgroups : [];
        var selectedSet = {};
        for (var i = 0; i < selected.length; i++) { selectedSet[selected[i]] = true; }

        var h = '';
        for (var j = 0; j < available.length; j++) {
            var sg = available[j];
            var isChecked = !!selectedSet[sg];
            h += '<div class="dr-pin-question-row">';
            h += '<label class="dr-flex dr-items-center dr-gap-8" style="cursor:pointer;flex:1;min-width:0">';
            h += '<input type="checkbox"' + (isChecked ? ' checked' : '') + ' data-sg="' + escAttr(sg) + '" onchange="DRApp.toggleDestSubgroup(this.dataset.sg)">';
            h += '<span class="dr-text-sm">' + escHtml(sg) + '</span>';
            h += '</label></div>';
        }
        $('#drDestSubgroupsList').html(h);
    }

    function toggleDestSubgroup(sgName) {
        if (!window._editingDest) return;
        if (!window._editingDest.subgroups) window._editingDest.subgroups = [];
        var idx = window._editingDest.subgroups.indexOf(sgName);
        if (idx >= 0) {
            window._editingDest.subgroups.splice(idx, 1);
        } else {
            window._editingDest.subgroups.push(sgName);
        }
        window._editingDest.subgroups.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
    }

    function cancelDestEdit() {
        $('#drDestList').html(renderDestList(editingStation.destinations));
    }

    function saveDest() {
        var dest = window._editingDest;
        dest.displayName = $('#drDestDisplayName').val();
        dest.orgId = parseInt($('#drDestOrgId').val()) || 0;
        dest.orgName = $('#drDestOrgName').val();
        dest.subgroup = $('#drDestSubgroup').val();
        dest.capacity = parseInt($('#drDestCapacity').val()) || 0;
        dest.capType = $('#drDestCapType').val();

        if (!dest.orgId) {
            showToast('Please select a destination involvement', 'warning');
            return;
        }

        if (editingDestIdx >= 0) {
            editingStation.destinations[editingDestIdx] = dest;
        } else {
            editingStation.destinations.push(dest);
        }

        editingDestIdx = -1;
        $('#drDestList').html(renderDestList(editingStation.destinations));
    }

    function clearStationSource() {
        editingStation.sourceOrgId = 0;
        editingStation.sourceOrgName = '';
        renderStationEditor();
    }

    function updateStationName(value) {
        if (editingStation) editingStation.name = value;
    }

    function clearDestOrg() {
        window._editingDest.orgId = 0;
        window._editingDest.orgName = '';
        window._editingDest.subgroups = [];
        $('#drDestOrgId').val(0);
        $('#drDestOrgName').val('');
        $('#drDestOrgSelected').replaceWith(
            '<div class="dr-search-box"><input type="text" class="dr-input" id="drDestOrgSearch" placeholder="Search involvements..." oninput="DRApp.searchDestOrg(this.value)">' +
            '<div class="dr-search-results" id="drDestOrgResults" style="display:none"></div></div>'
        );
        $('#drDestSubgroupsSection').hide();
        $('#drDestSubgroupsList').html('<div class="dr-text-center dr-text-muted dr-text-sm" style="padding:8px">Select an involvement first</div>');
    }

    // Involvement search for source/dest
    var searchTimer = null;

    function searchSource(term) {
        clearTimeout(searchTimer);
        if (term.length < 2) { $('#drSourceResults').hide(); return; }
        searchTimer = setTimeout(function() {
            ajax('search_involvements', {search_term: term}, function(resp) {
                if (resp.success) {
                    renderSearchResults('#drSourceResults', resp.involvements, 'source');
                }
            });
        }, 300);
    }

    function searchDestOrg(term) {
        clearTimeout(searchTimer);
        if (term.length < 2) { $('#drDestOrgResults').hide(); return; }
        searchTimer = setTimeout(function() {
            ajax('search_involvements', {search_term: term}, function(resp) {
                if (resp.success) {
                    renderSearchResults('#drDestOrgResults', resp.involvements, 'dest');
                }
            });
        }, 300);
    }

    function renderSearchResults(selector, involvements, mode) {
        var $el = $(selector);
        if (involvements.length === 0) {
            $el.html('<div style="padding:12px;color:#999">No results</div>').show();
            return;
        }
        // For dest mode, mark orgs already in this station's destinations
        var existing = {};
        if (mode === 'dest' && editingStation && editingStation.destinations) {
            for (var x = 0; x < editingStation.destinations.length; x++) {
                var d = editingStation.destinations[x];
                // Skip the dest currently being edited so its own org stays selectable
                if (editingDestIdx >= 0 && x === editingDestIdx) continue;
                if (d.orgId) existing[d.orgId] = true;
            }
        }
        // Sort: selectable first, already-added last (so clickable items are always visible)
        var sortedInvs = involvements.slice().sort(function(a, b) {
            var aAdded = !!existing[a.orgId] ? 1 : 0;
            var bAdded = !!existing[b.orgId] ? 1 : 0;
            return aAdded - bAdded;
        });
        var h = '';
        for (var i = 0; i < sortedInvs.length; i++) {
            var inv = sortedInvs[i];
            var alreadyAdded = !!existing[inv.orgId];
            if (alreadyAdded) {
                h += '<div class="dr-search-result-item" style="opacity:0.5;cursor:not-allowed">';
                h += '<div><div class="dr-search-result-name">' + escHtml(inv.orgName) + ' <span class="dr-text-sm dr-text-muted">(already added)</span></div>';
                h += '<div class="dr-search-result-meta">#' + inv.orgId + ' &middot; ' + escHtml(inv.programName) + ' &rarr; ' + escHtml(inv.divisionName) + ' &middot; ' + inv.memberCount + ' members</div>';
                h += '</div></div>';
            } else {
                h += '<div class="dr-search-result-item" data-mode="' + escAttr(mode) + '" data-org-id="' + inv.orgId + '" data-org-name="' + escAttr(inv.orgName) + '" onclick="DRApp.selectOrgFromEl(this)">';
                h += '<div><div class="dr-search-result-name">' + escHtml(inv.orgName) + '</div>';
                h += '<div class="dr-search-result-meta">#' + inv.orgId + ' &middot; ' + escHtml(inv.programName) + ' &rarr; ' + escHtml(inv.divisionName) + ' &middot; ' + inv.memberCount + ' members</div>';
                h += '</div></div>';
            }
        }
        $el.html(h).show();
    }

    function selectOrgFromEl(el) {
        selectOrg(el.getAttribute('data-mode'), parseInt(el.getAttribute('data-org-id'), 10), el.getAttribute('data-org-name') || '');
    }

    function selectOrg(mode, orgId, orgName) {
        if (mode === 'source') {
            editingStation.sourceOrgId = orgId;
            editingStation.sourceOrgName = orgName;
            renderStationEditor();
        } else {
            window._editingDest.orgId = orgId;
            window._editingDest.orgName = orgName;
            window._editingDest.subgroups = [];
            $('#drDestOrgId').val(orgId);
            $('#drDestOrgName').val(orgName);
            $('#drDestOrgSearch').closest('.dr-search-box').replaceWith(
                '<div class="dr-selected-org" id="drDestOrgSelected"><span>' + escHtml(orgName) + ' (#' + orgId + ')</span>' +
                '<span class="dr-remove-btn" onclick="DRApp.clearDestOrg()"><i class="bi bi-x-lg"></i></span></div>'
            );
            // Show subgroups section and load subgroups
            $('#drDestSubgroupsSection').show();
            loadDestSubgroups(orgId);
        }
    }

    function saveStation() {
        var st = editingStation;
        st.name = $('#drStationName').val();
        if (!st.name) {
            showToast('Station name is required', 'warning');
            return;
        }

        if (editingStationIdx >= 0) {
            state.currentScenario.stations[editingStationIdx] = st;
        } else {
            state.currentScenario.stations.push(st);
        }
        hideModal('drStationModal');
        $('#drStationList').html(renderStationsList(state.currentScenario.stations));
    }

    function toggleTestMode(checked) {
        if (state.currentScenario) {
            state.currentScenario.isTestMode = !!checked;
        }
        // Re-render so the danger button appears/disappears
        renderAdminEdit($('#drApp'));
    }

    function clearAssignmentLog() {
        if (!state.currentScenario || !state.currentScenario.id) {
            showToast('Save the scenario first', 'warning');
            return;
        }
        if (!confirm('Empty the assignment log for "' + state.currentScenario.name + '"?\\n\\nThis only affects the audit log. People remain assigned to their rooms.')) return;
        ajax('clear_assignment_log', {scenario_id: state.currentScenario.id}, function(resp) {
            if (resp.success) {
                showToast('Log cleared.', 'success');
            } else {
                showToast('Error: ' + (resp.message || 'unknown'), 'danger');
            }
        });
    }

    function resetTestAssignments() {
        if (!state.currentScenario || !state.currentScenario.id) {
            showToast('Save the scenario first', 'warning');
            return;
        }
        if (!state.currentScenario.isTestMode) {
            showToast('Enable Test Mode first', 'warning');
            return;
        }
        ajax('reset_test_assignments_preview', {scenario_id: state.currentScenario.id}, function(resp) {
            if (!resp.success) {
                showToast('Error: ' + (resp.message || 'unknown'), 'danger');
                return;
            }
            if (resp.totalAssigned === 0) {
                var none = 'No script-assigned people are currently in any scenario destination.';
                if (resp.logTotal) {
                    none += '\\n\\nThe log has ' + resp.logTotal + ' entry(ies), but ';
                    var notes = [];
                    if (resp.skippedStaleOrgs) notes.push(resp.skippedStaleOrgs + ' point to orgs no longer in this scenario');
                    if (resp.skippedAlreadyRemoved) notes.push(resp.skippedAlreadyRemoved + ' are already cleaned up');
                    none += notes.join(' and ') + '.';
                    none += '\\n\\nUse "Clear Assignment Log" to wipe the stale entries.';
                }
                alert(none);
                return;
            }
            var lines = ['You are about to remove ' + resp.totalAssigned + ' person-assignment(s) across ' + (resp.destinations || []).length + ' room(s):'];
            for (var i = 0; i < (resp.destinations || []).length; i++) {
                var d = resp.destinations[i];
                lines.push('  \\u2022 ' + (d.displayName || d.orgName) + ' (' + d.stationName + '): ' + d.count + ' person(s)');
                // List individual names so volunteers can spot anyone unexpected
                if (d.people && d.people.length > 0) {
                    var names = [];
                    for (var pi = 0; pi < d.people.length; pi++) {
                        names.push(d.people[pi].name || ('PeopleId ' + d.people[pi].peopleId));
                    }
                    lines.push('       ' + names.join(', '));
                }
            }
            // Show filter stats so stale-log situations are visible
            if (resp.skippedStaleOrgs || resp.skippedAlreadyRemoved) {
                lines.push('');
                var statsParts = [];
                if (resp.skippedStaleOrgs) statsParts.push(resp.skippedStaleOrgs + ' stale log entry(ies) for orgs no longer in this scenario');
                if (resp.skippedAlreadyRemoved) statsParts.push(resp.skippedAlreadyRemoved + ' entry(ies) already cleaned up');
                lines.push('(Filtered out: ' + statsParts.join(', ') + '.)');
            }
            lines.push('');
            lines.push('A backup will be saved before any changes are made.');
            lines.push('');
            lines.push('To confirm, type the EXACT scenario name below:');
            lines.push('  ' + resp.scenarioName);
            var typed = prompt(lines.join('\\n'), '');
            if (typed === null) return;
            if (typed.trim() !== (resp.scenarioName || '').trim()) {
                showToast('Name did not match. Reset cancelled.', 'warning');
                return;
            }
            ajax('reset_test_assignments', {scenario_id: state.currentScenario.id, confirm_name: typed}, function(r2) {
                if (r2.success) {
                    showToast(r2.message || 'Reset complete.', 'success');
                    if (r2.errors && r2.errors.length > 0) {
                        console.warn('[DR] Reset partial errors:', r2.errors);
                    }
                } else {
                    showToast('Error: ' + (r2.message || 'unknown'), 'danger');
                }
            });
        });
    }

    function saveScenario() {
        var sc = state.currentScenario;
        sc.name = $('#drScenarioName').val();
        sc.friendFinderProgramId = $('#drFFProgram').val() || '';
        sc.friendFinderDivisionId = $('#drFFDivision').val() || '';
        sc.isTestMode = $('#drIsTestMode').is(':checked');
        if (!sc.name) {
            showToast('Scenario name is required', 'warning');
            return;
        }
        ajax('save_scenario', {scenario_json: JSON.stringify(sc)}, function(resp) {
            if (resp.success) {
                showToast('Scenario saved!', 'success');
                state.currentScenario = resp.scenario;
                state.mode = 'admin';
                render();
            } else {
                showToast('Error: ' + resp.message, 'danger');
            }
        });
    }

    // ================================================================
    // STATION MODE - Scenario Picker
    // ================================================================
    function renderStationPickScenario($app) {
        var h = '<div class="dr-back-btn" onclick="DRApp.goLanding()"><i class="bi bi-arrow-left"></i> Back</div>';
        h += '<h2 class="dr-mb-16 dr-text-center">Select Event</h2>';
        h += '<div class="dr-picker-list" id="drScenarioPicker"><div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Loading...</div></div>';
        $app.html(h);

        ajax('load_scenarios', {}, function(resp) {
            if (resp.success) {
                state.scenarios = resp.scenarios || [];
                var h2 = '';
                if (state.scenarios.length === 0) {
                    h2 = '<div class="dr-empty"><i class="bi bi-inbox"></i><div>No scenarios found</div><div class="dr-text-sm">Create one in Admin Setup first</div></div>';
                } else {
                    for (var i = 0; i < state.scenarios.length; i++) {
                        var s = state.scenarios[i];
                        h2 += '<div class="dr-picker-btn" onclick="DRApp.pickScenario(\\'' + s.id + '\\')">';
                        h2 += '<i class="bi bi-calendar-event"></i> ' + escHtml(s.name);
                        h2 += '</div>';
                    }
                }
                $('#drScenarioPicker').html(h2);
            }
        });
    }

    function pickScenario(id) {
        for (var i = 0; i < state.scenarios.length; i++) {
            if (state.scenarios[i].id === id) {
                state.currentScenario = state.scenarios[i];
                state.mode = 'station_pick_station';
                render();
                return;
            }
        }
    }

    // ================================================================
    // STATION MODE - Station Picker
    // ================================================================
    function renderStationPickStation($app) {
        var sc = state.currentScenario;
        var h = '<div class="dr-back-btn" onclick="DRApp.navigate(\\'station_pick_scenario\\')"><i class="bi bi-arrow-left"></i> Back</div>';
        h += '<h2 class="dr-mb-8 dr-text-center">' + escHtml(sc.name) + '</h2>';
        h += '<p class="dr-text-center dr-text-muted dr-mb-16">Select your station</p>';
        h += '<div class="dr-picker-list">';
        var stations = sc.stations || [];
        for (var i = 0; i < stations.length; i++) {
            var st = stations[i];
            h += '<div class="dr-picker-btn" onclick="DRApp.pickStation(\\'' + st.id + '\\')">';
            h += '<i class="bi bi-person-workspace"></i> ' + escHtml(st.name);
            h += '</div>';
        }
        h += '</div>';
        $app.html(h);
    }

    function pickStation(id) {
        var stations = state.currentScenario.stations || [];
        for (var i = 0; i < stations.length; i++) {
            if (stations[i].id === id) {
                state.currentStation = stations[i];
                state.mode = 'station_pick_volunteer';
                render();
                return;
            }
        }
    }

    // ================================================================
    // STATION MODE - Volunteer Name
    // ================================================================
    function renderStationPickVolunteer($app) {
        var h = '<div class="dr-back-btn" onclick="DRApp.navigate(\\'station_pick_station\\')"><i class="bi bi-arrow-left"></i> Back</div>';
        h += '<div style="max-width:500px;margin:60px auto;text-align:center">';
        h += '<h2 class="dr-mb-8">Your Name</h2>';
        h += '<p class="dr-text-muted dr-mb-24">Enter your name so assignments are tracked</p>';
        h += '<input type="text" class="dr-input dr-mb-16" id="drVolunteerName" placeholder="Your name..." style="text-align:center;font-size:22px">';
        h += '<button class="dr-btn dr-btn-success dr-btn-lg dr-btn-block" onclick="DRApp.startStation()"><i class="bi bi-arrow-right"></i> Start Registering</button>';
        h += '</div>';
        $app.html(h);
        $('#drVolunteerName').focus();
    }

    function startStation() {
        var name = $('#drVolunteerName').val().trim();
        if (!name) {
            showToast('Please enter your name', 'warning');
            return;
        }
        state.volunteerName = name;
        state.mode = 'station_main';
        render();
        // Auto-enter fullscreen when station starts
        if (!state.fullscreen) { toggleFullscreen(); }
    }

    function toggleFullscreen() {
        state.fullscreen = !state.fullscreen;
        var root = document.querySelector('.dr-root');
        if (state.fullscreen) {
            if (root) root.classList.add('dr-fullscreen');
            // Walk up from .dr-root and hide all sibling elements at each level
            var el = root;
            while (el && el !== document.body) {
                var parent = el.parentElement;
                if (parent) {
                    var siblings = parent.children;
                    for (var i = 0; i < siblings.length; i++) {
                        if (siblings[i] !== el && !siblings[i].closest('.dr-root') && siblings[i].tagName !== 'SCRIPT' && siblings[i].tagName !== 'STYLE' && siblings[i].tagName !== 'LINK') {
                            siblings[i].setAttribute('data-dr-hidden', 'true');
                            siblings[i].style.display = 'none';
                        }
                    }
                }
                el = parent;
            }
        } else {
            if (root) root.classList.remove('dr-fullscreen');
            var hidden = document.querySelectorAll('[data-dr-hidden]');
            for (var i = 0; i < hidden.length; i++) {
                hidden[i].style.display = '';
                hidden[i].removeAttribute('data-dr-hidden');
            }
        }
        // Update toggle button icon if present
        var btn = document.getElementById('drFullscreenBtn');
        if (btn) {
            btn.innerHTML = state.fullscreen ? '<i class="bi bi-fullscreen-exit"></i>' : '<i class="bi bi-arrows-fullscreen"></i>';
        }
    }

    function exitFullscreen() {
        if (state.fullscreen) { toggleFullscreen(); }
    }

    // ================================================================
    // STATION MODE - Main Registration View
    // ================================================================
    function renderStationMain($app) {
        var sc = state.currentScenario;
        var st = state.currentStation;

        // Header
        var h = '<div class="dr-station-header-bar">';
        h += '<div><div class="dr-header-title">' + escHtml(sc.name) + ' &rsaquo; ' + escHtml(st.name) + '</div>';
        h += '<div class="dr-header-subtitle">Volunteer: ' + escHtml(state.volunteerName) + '</div></div>';
        h += '<div class="dr-header-actions">';
        h += '<button class="dr-btn dr-btn-warning dr-btn-sm" onclick="DRApp.showFindFriend()"><i class="bi bi-search-heart"></i> Find Friend</button>';
        h += '<button id="drFullscreenBtn" class="dr-btn dr-btn-outline dr-btn-sm" style="color:white;border-color:rgba(255,255,255,0.3);background:transparent" onclick="DRApp.toggleFullscreen()" title="Toggle fullscreen">' + (state.fullscreen ? '<i class="bi bi-fullscreen-exit"></i>' : '<i class="bi bi-arrows-fullscreen"></i>') + '</button>';
        h += '</div></div>';

        // Connection banner
        h += '<div class="dr-conn-banner" id="drConnBanner"><i class="bi bi-wifi-off"></i> Connection issue - retrying...</div>';

        // Search bar with sync status
        h += '<div class="dr-search-bar">';
        h += '<div class="dr-status-item"><span class="dr-status-dot dr-green" id="drSyncDot"></span> <span id="drSyncText">Syncing...</span></div>';
        h += '<div class="dr-search-wrap"><input type="text" class="dr-input" id="drSearchInput" placeholder="Search registrants..." oninput="DRApp.filterPeople(this.value)"><button class="dr-search-clear" id="drSearchClear" onclick="DRApp.clearSearch()">&times;</button></div>';
        h += '<button class="dr-btn dr-btn-outline" onclick="DRApp.searchAllPeople()"><i class="bi bi-globe"></i> All People</button>';
        h += '</div>';
        h += '<div id="drLastAction" style="display:none;font-size:13px;color:#666;margin:-8px 0 8px 0"><i class="bi bi-clock"></i> <span id="drLastActionText"></span></div>';

        // Person info bar (populated dynamically on selection)
        h += '<div id="drPersonInfoBar"></div>';

        // Main layout
        h += '<div class="dr-main-layout">';

        // Left: People panel
        h += '<div class="dr-people-panel"><div class="dr-panel">';
        h += '<div class="dr-people-tabs">';
        h += '<button class="dr-people-tab dr-tab-active" id="drTabPending" onclick="DRApp.switchPeopleTab(\\'pending\\')"><i class="bi bi-hourglass-split"></i> Pending <span class="dr-tab-count" id="drTabPendingCount">-</span></button>';
        h += '<button class="dr-people-tab" id="drTabCompleted" onclick="DRApp.switchPeopleTab(\\'completed\\')"><i class="bi bi-check-circle"></i> Completed <span class="dr-tab-count" id="drTabCompletedCount">-</span></button>';
        h += '</div>';
        h += '<div class="dr-panel-body" id="drPeopleList" style="max-height:calc(100vh - 360px);overflow-y:auto"><div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Loading registrants...</div></div>';
        h += '</div></div>';

        // Right: Destinations panel
        h += '<div class="dr-dest-panel">';
        h += '<div class="dr-dest-station-switcher" style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">';
        h += '<div><span class="dr-dest-station-label">Station:</span> ';
        h += '<span class="dr-station-switcher-wrap"><span class="dr-station-switch-trigger" id="drStationSwitchTrigger" onclick="DRApp.showStationSwitcher()">' + escHtml(st.name) + ' <i class="bi bi-chevron-down"></i></span><span id="drStationSwitcherDropdown"></span></span></div>';
        h += '<div style="font-size:0.85em"><span class="dr-text-muted">Sort:</span> ';
        h += '<select id="drDestSort" onchange="DRApp.changeDestSort(this.value)" style="padding:2px 6px;border:1px solid var(--dr-border);border-radius:4px;background:white;font-size:0.95em">';
        var sortOpts = [['default','Added'],['name','A \\u2192 Z'],['available','Most available'],['fullest','Most full']];
        for (var so = 0; so < sortOpts.length; so++) {
            var sel = state.destSort === sortOpts[so][0] ? ' selected' : '';
            h += '<option value="' + sortOpts[so][0] + '"' + sel + '>' + sortOpts[so][1] + '</option>';
        }
        h += '</select></div>';
        h += '</div>';
        h += '<div id="drDestCards"><div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Loading...</div></div>';
        h += '</div>';

        h += '</div>'; // end main-layout

        // Floating assign action bar (shown when person + destination selected)
        h += '<div id="drAssignActionBar"></div>';

        // Find Friend modal
        h += '<div class="dr-modal-overlay" id="drFindFriendModal">';
        h += '<div class="dr-modal">';
        h += '<div class="dr-modal-header"><span>Find a Friend</span>';
        h += '<button class="dr-modal-close" onclick="DRApp.hideModal(\\'drFindFriendModal\\')">&times;</button></div>';
        h += '<div class="dr-modal-body">';
        h += '<input type="text" class="dr-input dr-mb-12" id="drFriendSearch" placeholder="Search by name..." onkeydown="if(event.key===\\'Enter\\'){event.preventDefault();DRApp.doFindFriend();}">';
        h += '<button class="dr-btn dr-btn-primary dr-btn-block dr-mb-16" onclick="DRApp.doFindFriend()"><i class="bi bi-search"></i> Search</button>';
        h += '<div id="drFriendResults"></div>';
        h += '</div></div></div>';

        // Walk-in search modal
        h += '<div class="dr-modal-overlay" id="drWalkinModal">';
        h += '<div class="dr-modal">';
        h += '<div class="dr-modal-header"><span>Search All People</span>';
        h += '<button class="dr-modal-close" onclick="DRApp.hideModal(\\'drWalkinModal\\')">&times;</button></div>';
        h += '<div class="dr-modal-body">';
        h += '<input type="text" class="dr-input dr-mb-12" id="drWalkinSearch" placeholder="Search by name...">';
        h += '<button class="dr-btn dr-btn-primary dr-btn-block dr-mb-16" onclick="DRApp.doWalkinSearch()"><i class="bi bi-search"></i> Search</button>';
        h += '<div id="drWalkinResults"></div>';
        h += '</div></div></div>';

        // Person detail modal
        h += '<div class="dr-modal-overlay" id="drPersonDetailModal">';
        h += '<div class="dr-modal">';
        h += '<div class="dr-modal-header"><span id="drPersonDetailTitle">Person Details</span>';
        h += '<button class="dr-modal-close" onclick="DRApp.hideModal(\\'drPersonDetailModal\\')">&times;</button></div>';
        h += '<div class="dr-modal-body" id="drPersonDetailBody"></div>';
        h += '</div></div>';

        $app.html(h);
        loadStationData();
    }

    function loadStationData() {
        ajax('load_station_data', {
            scenario_id: state.currentScenario.id,
            station_id: state.currentStation.id
        }, function(resp) {
            if (resp.success) {
                state.registrants = resp.registrants || [];
                state.destinations = resp.destinations || [];
                state.assignedIds = resp.assignedIds || {};
                state.peopleTab = 'pending';
                renderPeopleList();
                renderDestCards();
                updateStatusCounts();
                startPolling();
            } else {
                console.error('[DR] loadStationData failed:', resp.message);
                $('#drPeopleList').html('<div class="dr-empty"><i class="bi bi-exclamation-triangle"></i><div>Error loading registrants</div><div class="dr-text-sm">' + escHtml(resp.message) + '</div></div>');
                $('#drDestCards').html('<div class="dr-empty"><i class="bi bi-exclamation-triangle"></i><div>Error loading destinations</div><div class="dr-text-sm">' + escHtml(resp.message) + '</div></div>');
                showToast('Error loading data: ' + resp.message, 'danger');
            }
        });
    }

    function renderPeopleList(filter) {
        var $list = $('#drPeopleList');
        var people = state.registrants;
        var tab = state.peopleTab || 'pending';

        // Filter by search term
        if (filter) {
            var f = filter.toLowerCase();
            people = people.filter(function(p) {
                return (p.name || '').toLowerCase().indexOf(f) >= 0 ||
                       (p.firstName || '').toLowerCase().indexOf(f) >= 0 ||
                       (p.lastName || '').toLowerCase().indexOf(f) >= 0 ||
                       (p.nickName || '').toLowerCase().indexOf(f) >= 0;
            });
        }

        // Filter by tab
        people = people.filter(function(p) {
            var isAssigned = !!state.assignedIds[String(p.peopleId)];
            return tab === 'completed' ? isAssigned : !isAssigned;
        });

        if (people.length === 0) {
            var emptyMsg = tab === 'completed' ? 'No completed registrants' : 'No pending registrants';
            var emptyIcon = tab === 'completed' ? 'bi-check-circle' : 'bi-hourglass-split';
            $list.html('<div class="dr-empty"><i class="bi ' + emptyIcon + '"></i><div>' + emptyMsg + '</div></div>');
            return;
        }

        var h = '';
        for (var i = 0; i < people.length; i++) {
            var p = people[i];
            var isAssigned = !!state.assignedIds[String(p.peopleId)];
            var isSelected = state.selectedPerson && state.selectedPerson.peopleId === p.peopleId;
            var classes = 'dr-person-card';
            if (isSelected) classes += ' dr-selected';
            if (isAssigned) classes += ' dr-assigned';

            h += '<div class="' + classes + '" onclick="DRApp.selectPerson(' + p.peopleId + ')" data-pid="' + p.peopleId + '">';
            h += '<div><div class="dr-person-name">' + escHtml(p.name) + '</div>';
            h += '<div class="dr-person-meta">';
            if (p.age) h += 'Age ' + p.age + ' &middot; ';
            if (p.gender) h += p.gender;
            h += '</div></div>';

            if (isAssigned) {
                var info = state.assignedIds[String(p.peopleId)];
                h += '<div><span class="dr-person-badge dr-badge-assigned"><i class="bi bi-check-circle-fill"></i>&nbsp;' + escHtml(info.destDisplayName || info.destOrgName || 'Assigned') + '</span></div>';
            }
            h += '</div>';
        }
        $list.html(h);
    }

    function renderDestCards() {
        var $panel = $('#drDestCards');
        var dests = state.destinations;
        if (dests.length === 0) {
            $panel.html('<div class="dr-empty"><i class="bi bi-door-open"></i><div>No destinations configured</div></div>');
            return;
        }
        // Build display order with original-index preservation so colors and click handlers stay correct
        var displayOrder = [];
        for (var oi = 0; oi < dests.length; oi++) displayOrder.push(oi);
        var sortMode = state.destSort || 'default';
        if (sortMode === 'name') {
            displayOrder.sort(function(a, b) {
                var na = (dests[a].displayName || dests[a].orgName || '').toLowerCase();
                var nb = (dests[b].displayName || dests[b].orgName || '').toLowerCase();
                return na.localeCompare(nb);
            });
        } else if (sortMode === 'available') {
            displayOrder.sort(function(a, b) {
                var da = dests[a], db = dests[b];
                var availA = (da.capacity > 0) ? (da.capacity - (da.currentCount || 0)) : Number.POSITIVE_INFINITY;
                var availB = (db.capacity > 0) ? (db.capacity - (db.currentCount || 0)) : Number.POSITIVE_INFINITY;
                return availB - availA;
            });
        } else if (sortMode === 'fullest') {
            displayOrder.sort(function(a, b) {
                var da = dests[a], db = dests[b];
                var pctA = (da.capacity > 0) ? ((da.currentCount || 0) / da.capacity) : 0;
                var pctB = (db.capacity > 0) ? ((db.currentCount || 0) / db.capacity) : 0;
                return pctB - pctA;
            });
        }
        var h = '';
        for (var oo = 0; oo < displayOrder.length; oo++) {
            var i = displayOrder[oo];
            var d = dests[i];
            var colorClass = 'dr-dest-' + (i % 8);
            var pct = d.capacity > 0 ? Math.min(100, Math.round((d.currentCount / d.capacity) * 100)) : 0;
            var capClass = pct < 80 ? 'dr-cap-ok' : (pct < 100 ? 'dr-cap-warn' : 'dr-cap-full');
            var disabled = (d.capType === 'hard' && d.capacity > 0 && d.currentCount >= d.capacity);

            var isSelected = (state.selectedDestIdx === i);
            var selClass = isSelected ? ' dr-dest-selected dr-dest-expanded' : '';
            var disClass = disabled ? ' dr-dest-disabled' : '';

            // Collect names assigned to this destination
            var assignedNames = [];
            for (var pid in state.assignedIds) {
                var info = state.assignedIds[pid];
                if (info.destOrgId === d.orgId || (!info.destOrgId && (info.destOrgName === d.orgName))) {
                    assignedNames.push(info.personName || 'Unknown');
                }
            }
            assignedNames.sort();

            h += '<div class="dr-dest-card ' + colorClass + selClass + disClass + '" data-dest-id="' + d.destId + '" data-org-id="' + d.orgId + '" onclick="DRApp.selectDest(' + i + ')">';
            h += '<div class="dr-dest-card-header">';
            h += '<div class="dr-dest-card-name">' + escHtml(d.displayName || d.orgName) + '</div>';
            h += '<div class="dr-dest-card-count" id="drDestCount_' + d.orgId + '">' + d.currentCount + (d.capacity > 0 ? '/' + d.capacity : '') + '</div>';
            h += '</div>';
            if (d.capacity > 0) {
                h += '<div class="dr-dest-card-body">';
                h += '<div class="dr-cap-bar-wrap"><div class="dr-cap-bar ' + capClass + '" id="drCapBar_' + d.orgId + '" style="width:' + pct + '%"></div></div>';
                h += '</div>';
            }
            var hasSubgroups = d.subgroups && d.subgroups.length > 0;
            if (hasSubgroups || assignedNames.length > 0) {
                h += '<div class="dr-dest-card-detail">';
                if (hasSubgroups) {
                    d.subgroups.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
                    h += '<div class="dr-dest-subgroups">';
                    h += '<div class="dr-text-sm dr-text-muted dr-mb-4" style="font-weight:600">Subgroups:</div>';
                    var selSgs = state.selectedSubgroups || [];
                    var sgCounts = d.subgroupCounts || {};
                    for (var sg = 0; sg < d.subgroups.length; sg++) {
                        var sgName = d.subgroups[sg];
                        var sgActive = selSgs.indexOf(sgName) >= 0;
                        var sgCnt = sgCounts[sgName] !== undefined ? sgCounts[sgName] : '';
                        var sgCountLabel = sgCnt !== '' ? ' (' + sgCnt + ')' : '';
                        h += '<button class="dr-subgroup-btn' + (sgActive ? ' dr-sg-active' : '') + '" data-dest="' + i + '" data-sg="' + escAttr(sgName) + '" onclick="event.stopPropagation(); DRApp.toggleSubgroup(this.dataset.sg)">';
                        h += '<i class="bi ' + (sgActive ? 'bi-check-circle-fill' : 'bi-circle') + '"></i> ' + escHtml(sgName) + sgCountLabel;
                        h += '</button>';
                    }
                    h += '</div>';
                }
                if (assignedNames.length > 0) {
                    h += '<div class="dr-dest-names">';
                    for (var n = 0; n < assignedNames.length; n++) {
                        h += '<div class="dr-dest-name-item">' + escHtml(assignedNames[n]) + '</div>';
                    }
                    h += '</div>';
                }
                h += '</div>';
            }
            h += '</div>';
        }
        $panel.html(h);
        renderAssignBar();
    }

    function changeDestSort(mode) {
        state.destSort = mode || 'default';
        try { localStorage.setItem('drDestSort', state.destSort); } catch(e) {}
        renderDestCards();
    }

    function selectDest(destIdx) {
        var dest = state.destinations[destIdx];
        if (!dest) return;

        // Check hard capacity
        var disabled = (dest.capType === 'hard' && dest.capacity > 0 && dest.currentCount >= dest.capacity);
        if (disabled) {
            showToast(escHtml(dest.displayName || dest.orgName) + ' is full (hard cap reached)', 'danger');
            return;
        }

        // Toggle selection
        if (state.selectedDestIdx === destIdx) {
            state.selectedDestIdx = null;
            state.selectedSubgroups = [];
        } else {
            state.selectedDestIdx = destIdx;
            state.selectedSubgroups = [];
            if (!state.selectedPerson) {
                showToast('Now select a person', 'info');
            }
        }
        renderDestCards();
    }

    function renderAssignBar() {
        var $bar = $('#drAssignActionBar');
        if (!$bar.length) return;

        if (state.selectedPerson && state.selectedDestIdx !== null) {
            var person = state.selectedPerson;
            var dest = state.destinations[state.selectedDestIdx];
            if (!dest) { $bar.html(''); return; }

            var sgText = '';
            if (state.selectedSubgroups && state.selectedSubgroups.length > 0) {
                sgText = ' > ' + state.selectedSubgroups.join(', ');
            }
            var h = '<div class="dr-assign-action-bar">';
            h += '<div class="dr-assign-bar-info">';
            h += '<i class="bi bi-person-fill"></i> ' + escHtml(person.preferredName || person.firstName || '') + ' ' + escHtml(person.lastName || '');
            h += '<span class="dr-assign-bar-arrow">&rarr;</span>';
            h += '<i class="bi bi-door-open-fill"></i> ' + escHtml(dest.orgName || dest.displayName) + escHtml(sgText);
            h += '</div>';
            h += '<button class="dr-assign-bar-btn" onclick="DRApp.assignToDest()">';
            h += '<i class="bi bi-person-plus-fill"></i> ASSIGN';
            h += '</button>';
            h += '</div>';
            $bar.html(h);
        } else {
            $bar.html('');
        }
    }

    function selectPerson(pid) {
        // Check if person is assigned - show undo option
        if (state.assignedIds[String(pid)]) {
            var info = state.assignedIds[String(pid)];
            if (confirm('This person is already assigned to ' + (info.destDisplayName || info.destOrgName) + '. Would you like to undo this assignment?')) {
                undoAssignment(pid, info);
            }
            return;
        }

        var person = null;
        for (var i = 0; i < state.registrants.length; i++) {
            if (state.registrants[i].peopleId === pid) {
                person = state.registrants[i];
                break;
            }
        }
        // Also check walk-in selections
        if (!person && state._walkinPeople) {
            for (var j = 0; j < state._walkinPeople.length; j++) {
                if (state._walkinPeople[j].peopleId === pid) {
                    person = state._walkinPeople[j];
                    break;
                }
            }
        }

        if (state.selectedPerson && state.selectedPerson.peopleId === pid) {
            // Deselect
            state.selectedPerson = null;
            state.selectedDestIdx = null;
            state.personDetail = null;
            $('#drPersonInfoBar').html('');
            renderAssignBar();
            renderDestCards();
        } else {
            state.selectedPerson = person;
            state.personDetail = null;
            $('#drPersonInfoBar').html('');
            if (state.selectedDestIdx === null) {
                showToast('Now tap a destination', 'info');
            }
            renderAssignBar();
            // Fetch full detail including emergency/medical/reg answers
            if (person) {
                var sourceOrgId = (state.currentStation && state.currentStation.sourceOrgId) ? state.currentStation.sourceOrgId : 0;
                ajax('get_person_detail', {
                    people_id: person.peopleId,
                    source_org_id: sourceOrgId
                }, function(resp) {
                    if (resp.success && state.selectedPerson && state.selectedPerson.peopleId === person.peopleId) {
                        state.personDetail = resp;
                        renderPersonInfoBar();
                    }
                });
            }
        }
        renderPeopleList($('#drSearchInput').val());
    }

    function dismissPersonInfoBar() {
        state.selectedPerson = null;
        state.selectedDestIdx = null;
        state.personDetail = null;
        $('#drPersonInfoBar').html('');
        renderAssignBar();
        renderDestCards();
        renderPeopleList($('#drSearchInput').val());
    }

    function toggleRegAnswers() {
        $('#drInfoAnswersContent').toggleClass('dr-show');
        var $icon = $('#drInfoAnswersToggleIcon');
        if ($('#drInfoAnswersContent').hasClass('dr-show')) {
            $icon.removeClass('bi-chevron-right').addClass('bi-chevron-down');
        } else {
            $icon.removeClass('bi-chevron-down').addClass('bi-chevron-right');
        }
    }

    function renderPersonInfoBar() {
        var detail = state.personDetail;
        if (!detail) { $('#drPersonInfoBar').html(''); return; }

        var p = detail.person || {};
        var em = detail.emergency || {};
        var answers = detail.regAnswers || [];
        var family = detail.family || [];

        var ageGender = '';
        if (p.age) ageGender += 'Age ' + escHtml(p.age);
        if (p.age && p.gender) ageGender += ', ';
        if (p.gender) ageGender += escHtml(p.gender);

        var h = '<div class="dr-info-bar">';

        // Header
        h += '<div class="dr-info-bar-header">';
        h += '<span><i class="bi bi-person-check-fill"></i> SELECTED: ' + escHtml(p.name) + (ageGender ? ' (' + ageGender + ')' : '') + '</span>';
        h += '<button class="dr-info-bar-dismiss" onclick="DRApp.dismissPersonInfoBar()" title="Dismiss"><i class="bi bi-x-lg"></i></button>';
        h += '</div>';

        // Body
        h += '<div class="dr-info-bar-body">';

        // Emergency contact
        var hasEmContact = em.contactName && em.contactName.length > 0;
        h += '<div class="dr-info-section">';
        h += '<span class="dr-info-section-icon"><i class="bi bi-telephone-fill" style="color:var(--dr-danger)"></i></span>';
        h += '<span class="dr-info-label">Emergency:</span> ';
        if (hasEmContact) {
            h += '<span class="dr-info-value">' + escHtml(em.contactName);
            if (em.contactPhone) h += ' &bull; ' + escHtml(em.contactPhone);
            h += '</span>';
        } else {
            h += '<span class="dr-info-value" style="color:var(--dr-gray);font-style:italic">None on file</span>';
        }
        h += '</div>';

        // Medical
        var hasMedical = (em.medical && em.medical.length > 0) || (em.allergies && em.allergies.length > 0);
        h += '<div class="dr-info-section">';
        h += '<span class="dr-info-section-icon"><i class="bi bi-heart-pulse-fill" style="color:' + (hasMedical ? 'var(--dr-danger)' : 'var(--dr-secondary)') + '"></i></span>';
        h += '<span class="dr-info-label">Medical:</span> ';
        if (hasMedical) {
            var medParts = [];
            if (em.allergies) medParts.push(escHtml(em.allergies));
            if (em.medical) medParts.push(escHtml(em.medical));
            h += '<span class="dr-info-medical-alert">' + medParts.join(', ') + '</span>';
        } else {
            h += '<span class="dr-info-medical-clear">None reported</span>';
        }
        h += '</div>';

        // OTC Meds
        if (em.otcMeds && em.otcMeds.length > 0) {
            h += '<div class="dr-info-section">';
            h += '<span class="dr-info-section-icon"><i class="bi bi-capsule" style="color:var(--dr-primary)"></i></span>';
            h += '<span class="dr-info-label">OTC Meds:</span> ';
            h += '<span class="dr-info-value">' + escHtml(em.otcMeds.join(', ')) + '</span>';
            h += '</div>';
        }

        // Family
        if (family.length > 0) {
            h += '<div class="dr-info-section">';
            h += '<span class="dr-info-section-icon"><i class="bi bi-people-fill" style="color:var(--dr-primary)"></i></span>';
            h += '<span class="dr-info-label">Family:</span> ';
            h += '<span class="dr-info-family">';
            for (var fi = 0; fi < family.length; fi++) {
                if (fi > 0) h += '<span class="dr-info-family-sep"> &bull; </span>';
                h += escHtml(family[fi].name);
                var famMeta = '';
                if (family[fi].position) famMeta = family[fi].position;
                else if (family[fi].age) famMeta = 'Age ' + family[fi].age;
                if (famMeta) h += ' <span style="color:var(--dr-gray);font-size:13px">(' + escHtml(famMeta) + ')</span>';
            }
            h += '</span></div>';
        }

        // Pinned registration answers (shown inline in body)
        var pinnedQs = (state.currentStation && state.currentStation.pinnedQuestions) ? state.currentStation.pinnedQuestions : [];
        var pinnedSet = {};
        for (var pi = 0; pi < pinnedQs.length; pi++) {
            pinnedSet[pinnedQs[pi].question] = pinnedQs[pi].shortLabel || '';
        }
        var pinnedAnswers = [];
        var otherAnswers = [];
        for (var ai = 0; ai < answers.length; ai++) {
            if (pinnedSet.hasOwnProperty(answers[ai].question)) {
                pinnedAnswers.push({ question: answers[ai].question, answer: answers[ai].answer, shortLabel: pinnedSet[answers[ai].question] });
            } else {
                otherAnswers.push(answers[ai]);
            }
        }
        for (var pa = 0; pa < pinnedAnswers.length; pa++) {
            var dispLabel = pinnedAnswers[pa].shortLabel || pinnedAnswers[pa].question;
            h += '<div class="dr-info-section">';
            h += '<span class="dr-info-section-icon"><i class="bi bi-pin-fill" style="color:var(--dr-primary)"></i></span>';
            h += '<span class="dr-info-label">' + escHtml(dispLabel) + ':</span> ';
            h += '<span class="dr-info-value">' + escHtml(pinnedAnswers[pa].answer || '\\u2014') + '</span>';
            h += '</div>';
        }

        h += '</div>'; // end body

        // Registration answers (collapsible footer) - only non-pinned
        if (otherAnswers.length > 0) {
            h += '<div class="dr-info-bar-footer">';
            h += '<button class="dr-info-answers-toggle" onclick="DRApp.toggleRegAnswers()">';
            h += '<i class="bi bi-chevron-right" id="drInfoAnswersToggleIcon"></i> Registration Answers (' + otherAnswers.length + ')';
            h += '</button>';
            h += '<div class="dr-info-answers-content" id="drInfoAnswersContent">';
            for (var oi = 0; oi < otherAnswers.length; oi++) {
                h += '<div class="dr-info-answer-row">';
                h += '<span class="dr-info-answer-q">' + escHtml(otherAnswers[oi].question) + ':</span>';
                h += '<span class="dr-info-answer-a">' + escHtml(otherAnswers[oi].answer || '\\u2014') + '</span>';
                h += '</div>';
            }
            h += '</div></div>';
        }

        h += '</div>'; // end info-bar
        $('#drPersonInfoBar').html(h);
    }

    function assignToDest() {
        if (!state.selectedPerson) {
            showToast('Select a person first', 'warning');
            return;
        }
        if (state.selectedDestIdx === null) {
            showToast('Select a destination first', 'warning');
            return;
        }

        var dest = state.destinations[state.selectedDestIdx];
        var person = state.selectedPerson;

        // Disable the assign bar button during request
        var $barBtn = $('.dr-assign-bar-btn');
        $barBtn.prop('disabled', true).html('<div class="dr-spinner" style="width:20px;height:20px;border-width:2px;margin:0 auto"></div>');

        // Collect subgroups: selected toggles or legacy single subgroup
        var subgroupsToSend = (state.selectedSubgroups && state.selectedSubgroups.length > 0) ? state.selectedSubgroups.join('|') : (dest.subgroup || '');

        ajax('process_assignment', {
            people_id: person.peopleId,
            dest_org_id: dest.orgId,
            dest_subgroup: subgroupsToSend,
            capacity: dest.capacity || 0,
            cap_type: dest.capType || 'soft',
            scenario_id: state.currentScenario.id,
            station_id: state.currentStation.id,
            station_name: state.currentStation.name,
            volunteer_name: state.volunteerName,
            dest_display_name: dest.displayName || dest.orgName,
            dest_org_name: dest.orgName || ''
        }, function(resp) {
            if (resp.success) {
                // Update local state
                dest.currentCount = resp.currentCount;
                if (resp.subgroupCounts) dest.subgroupCounts = resp.subgroupCounts;
                state.assignedIds[String(person.peopleId)] = {
                    destOrgId: dest.orgId,
                    destOrgName: dest.orgName,
                    stationName: state.currentStation.name,
                    destDisplayName: dest.displayName || dest.orgName,
                    personName: person.name || (person.firstName + ' ' + person.lastName),
                    destSubgroup: subgroupsToSend
                };
                state.selectedPerson = null;
                state.selectedDestIdx = null;
                state.selectedSubgroups = [];
                state.personDetail = null;
                $('#drPersonInfoBar').html('');
                renderAssignBar();

                renderPeopleList($('#drSearchInput').val());
                renderDestCards();
                updateStatusCounts();

                var sgLabel = subgroupsToSend ? ' > ' + subgroupsToSend.replace(/\\|/g, ', ') : '';
                var msg = resp.personName + ' assigned to ' + (dest.displayName || dest.orgName) + sgLabel;
                if (resp.overCap) {
                    showToast('WARNING: Over capacity! ' + msg, 'warning');
                } else {
                    showToast(msg, 'success');
                }

                // Show last action
                $('#drLastAction').show();
                $('#drLastActionText').text(msg);
            } else if (resp.blocked) {
                showToast(resp.message, 'danger');
                dest.currentCount = resp.currentCount;
                renderDestCards();
            } else {
                showToast('Error: ' + resp.message, 'danger');
            }
            $barBtn.prop('disabled', false).html('<i class="bi bi-person-plus-fill"></i> ASSIGN');
        });
    }

    function toggleSubgroup(sgName) {
        if (!state.selectedSubgroups) state.selectedSubgroups = [];
        var idx = state.selectedSubgroups.indexOf(sgName);
        if (idx >= 0) {
            state.selectedSubgroups.splice(idx, 1);
        } else {
            state.selectedSubgroups.push(sgName);
        }
        renderDestCards();
        renderAssignBar();
    }

    function undoAssignment(pid, info) {
        // Use destOrgId directly from info (stored during assignment)
        var destOrgId = info.destOrgId || 0;
        var destSubgroup = info.destSubgroup || '';
        var destDisplayName = info.destDisplayName || info.destOrgName || '';
        var destOrgName = info.destOrgName || '';

        // Fallback: try to find org by name if destOrgId not in info
        if (!destOrgId) {
            for (var i = 0; i < state.destinations.length; i++) {
                if (state.destinations[i].displayName === destDisplayName || state.destinations[i].orgName === destOrgName) {
                    destOrgId = state.destinations[i].orgId;
                    if (!destSubgroup) destSubgroup = state.destinations[i].subgroup || '';
                    destOrgName = state.destinations[i].orgName;
                    break;
                }
            }
        }

        if (!destOrgId) {
            showToast('Could not find destination to undo', 'danger');
            return;
        }

        ajax('undo_assignment', {
            people_id: pid,
            dest_org_id: destOrgId,
            dest_subgroup: destSubgroup,
            scenario_id: state.currentScenario.id,
            station_id: state.currentStation.id,
            station_name: state.currentStation.name,
            volunteer_name: state.volunteerName,
            dest_display_name: destDisplayName,
            dest_org_name: destOrgName
        }, function(resp) {
            if (resp.success) {
                delete state.assignedIds[String(pid)];
                // Update dest count
                for (var i = 0; i < state.destinations.length; i++) {
                    if (state.destinations[i].orgId === destOrgId) {
                        state.destinations[i].currentCount = resp.currentCount;
                        if (resp.subgroupCounts) state.destinations[i].subgroupCounts = resp.subgroupCounts;
                        break;
                    }
                }
                renderPeopleList($('#drSearchInput').val());
                renderDestCards();
                updateStatusCounts();
                showToast(resp.message, 'success');
            } else {
                showToast('Error: ' + resp.message, 'danger');
            }
        });
    }

    function updateStatusCounts() {
        var total = state.registrants.length;
        var assigned = 0;
        for (var i = 0; i < state.registrants.length; i++) {
            if (state.assignedIds[String(state.registrants[i].peopleId)]) {
                assigned++;
            }
        }
        var pending = total - assigned;
        $('#drTabPendingCount').text(pending);
        $('#drTabCompletedCount').text(assigned);
    }

    function filterPeople(term) {
        state.searchTerm = term;
        $('#drSearchClear').toggle(!!term);
        renderPeopleList(term);
    }

    function clearSearch() {
        $('#drSearchInput').val('');
        filterPeople('');
    }

    function switchPeopleTab(tab) {
        state.peopleTab = tab;
        $('#drTabPending').toggleClass('dr-tab-active', tab === 'pending');
        $('#drTabCompleted').toggleClass('dr-tab-active', tab === 'completed');
        renderPeopleList($('#drSearchInput').val());
    }

    // ---- Polling ----
    function startPolling() {
        stopPolling();
        state.pollTimer = setInterval(function() {
            pollCounts();
            pollSync();
        }, 10000);
        $('#drSyncText').text('Connected');
        $('#drSyncDot').removeClass('dr-yellow').addClass('dr-green');
    }

    function stopPolling() {
        if (state.pollTimer) {
            clearInterval(state.pollTimer);
            state.pollTimer = null;
        }
    }

    function pollCounts() {
        var orgIds = [];
        for (var i = 0; i < state.destinations.length; i++) {
            orgIds.push(state.destinations[i].orgId);
        }
        if (orgIds.length === 0) return;

        ajax('get_destination_counts', {org_ids: orgIds.join(',')}, function(resp) {
            if (resp.success) {
                state.pollFailCount = 0;
                $('#drConnBanner').hide();
                $('#drSyncText').text('Connected');
                $('#drSyncDot').removeClass('dr-yellow').addClass('dr-green');

                var sgCountsResp = resp.subgroupCounts || {};
                var sgChanged = false;
                for (var i = 0; i < state.destinations.length; i++) {
                    var d = state.destinations[i];
                    var newCount = resp.counts[String(d.orgId)];
                    if (newCount !== undefined && newCount !== d.currentCount) {
                        d.currentCount = newCount;
                        // Update UI without full re-render
                        var countText = d.currentCount + (d.capacity > 0 ? '/' + d.capacity : '');
                        $('#drDestCount_' + d.orgId).text(countText);
                        if (d.capacity > 0) {
                            var pct = Math.min(100, Math.round((d.currentCount / d.capacity) * 100));
                            var capClass = pct < 80 ? 'dr-cap-ok' : (pct < 100 ? 'dr-cap-warn' : 'dr-cap-full');
                            $('#drCapBar_' + d.orgId).css('width', pct + '%').attr('class', 'dr-cap-bar ' + capClass);
                        }
                        if (d.capType === 'hard' && d.capacity > 0 && d.currentCount >= d.capacity) {
                            $('#drAssignBtn_' + d.orgId).prop('disabled', true);
                        } else {
                            $('#drAssignBtn_' + d.orgId).prop('disabled', false);
                        }
                    }
                    // Update subgroup counts
                    var newSgCounts = sgCountsResp[String(d.orgId)];
                    if (newSgCounts) {
                        d.subgroupCounts = newSgCounts;
                        sgChanged = true;
                    }
                }
                if (sgChanged) renderDestCards();
            } else {
                handlePollFailure();
            }
        });
    }

    function pollSync() {
        ajax('get_sync_status', {scenario_id: state.currentScenario.id}, function(resp) {
            if (resp.success) {
                var newAssigned = resp.assigned || {};
                var changed = false;
                // Check for new assignments from other stations
                for (var pid in newAssigned) {
                    if (!state.assignedIds[pid]) {
                        state.assignedIds[pid] = newAssigned[pid];
                        changed = true;
                    }
                }
                // Check for removed assignments
                for (var pid2 in state.assignedIds) {
                    if (!newAssigned[pid2]) {
                        delete state.assignedIds[pid2];
                        changed = true;
                    }
                }
                if (changed) {
                    renderPeopleList($('#drSearchInput').val());
                    renderDestCards();
                    updateStatusCounts();
                }
            }
        });
    }

    function handlePollFailure() {
        state.pollFailCount++;
        if (state.pollFailCount >= 2) {
            $('#drConnBanner').show();
            $('#drSyncText').text('Connection issue');
            $('#drSyncDot').removeClass('dr-green').addClass('dr-yellow');
        }
    }

    // ---- Find Friend ----
    function showFindFriend() { showModal('drFindFriendModal'); $('#drFriendSearch').val('').focus(); }

    function doFindFriend() {
        var term = $('#drFriendSearch').val().trim();
        if (!term) return;
        $('#drFriendResults').html('<div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Searching...</div>');

        ajax('find_friend', {search_term: term, scenario_id: state.currentScenario.id}, function(resp) {
            if (resp.success) {
                var results = resp.results || [];
                if (results.length === 0) {
                    $('#drFriendResults').html('<div class="dr-text-center dr-text-muted" style="padding:16px">No results found</div>');
                    return;
                }
                var h = '';
                for (var i = 0; i < results.length; i++) {
                    var r = results[i];
                    h += '<div class="dr-person-card" style="cursor:default">';
                    h += '<div><div class="dr-person-name">' + escHtml(r.name) + '</div>';
                    h += '<div class="dr-person-meta">';
                    if (r.age) h += 'Age ' + r.age + ' &middot; ';
                    if (r.gender) h += r.gender;
                    h += '</div>';
                    // Destination assignments (from this scenario)
                    if (r.destinations && r.destinations.length > 0) {
                        for (var j = 0; j < r.destinations.length; j++) {
                            h += '<span class="dr-person-badge dr-badge-assigned" style="margin-top:4px"><i class="bi bi-pin-map-fill"></i>&nbsp;' + escHtml(r.destinations[j].displayName) + ' (' + escHtml(r.destinations[j].stationName) + ')</span> ';
                        }
                    }
                    // Filtered involvements (from program/division filter)
                    if (r.involvements && r.involvements.length > 0) {
                        h += '<div style="margin-top:4px">';
                        for (var k = 0; k < r.involvements.length; k++) {
                            h += '<span class="dr-person-badge" style="margin-top:2px;background:#e8f4fd;color:#1565c0;border:1px solid #90caf9"><i class="bi bi-people-fill"></i>&nbsp;' + escHtml(r.involvements[k].orgName) + '</span> ';
                        }
                        h += '</div>';
                    }
                    // Show "not assigned" only if no destinations AND no involvements
                    if ((!r.destinations || r.destinations.length === 0) && (!r.involvements || r.involvements.length === 0)) {
                        h += '<span class="dr-text-sm dr-text-muted">Not assigned yet</span>';
                    }
                    h += '</div></div>';
                }
                $('#drFriendResults').html(h);
            } else {
                $('#drFriendResults').html('<div class="dr-text-center" style="padding:16px;color:red">' + resp.message + '</div>');
            }
        });
    }

    // ---- Walk-in Search ----
    function searchAllPeople() { showModal('drWalkinModal'); $('#drWalkinSearch').val('').focus(); }

    function doWalkinSearch() {
        var term = $('#drWalkinSearch').val().trim();
        if (term.length < 2) { showToast('Enter at least 2 characters', 'warning'); return; }
        $('#drWalkinResults').html('<div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Searching...</div>');

        ajax('search_all_people', {search_term: term}, function(resp) {
            if (resp.success) {
                state._walkinPeople = resp.people || [];
                var people = resp.people || [];
                if (people.length === 0) {
                    $('#drWalkinResults').html('<div class="dr-text-center dr-text-muted" style="padding:16px">No results</div>');
                    return;
                }
                var h = '';
                for (var i = 0; i < people.length; i++) {
                    var p = people[i];
                    var isAssigned = !!state.assignedIds[String(p.peopleId)];
                    h += '<div class="dr-person-card' + (isAssigned ? ' dr-assigned' : '') + '" onclick="DRApp.selectWalkin(' + p.peopleId + ')">';
                    h += '<div><div class="dr-person-name">' + escHtml(p.name) + '</div>';
                    h += '<div class="dr-person-meta">';
                    if (p.age) h += 'Age ' + p.age + ' &middot; ';
                    if (p.gender) h += p.gender;
                    h += '</div></div>';
                    if (isAssigned) {
                        h += '<span class="dr-person-badge dr-badge-assigned"><i class="bi bi-check-circle-fill"></i> Assigned</span>';
                    }
                    h += '</div>';
                }
                $('#drWalkinResults').html(h);
            }
        });
    }

    function selectWalkin(pid) {
        if (state.assignedIds[String(pid)]) {
            showToast('This person is already assigned', 'warning');
            return;
        }
        var person = null;
        if (state._walkinPeople) {
            for (var i = 0; i < state._walkinPeople.length; i++) {
                if (state._walkinPeople[i].peopleId === pid) {
                    person = state._walkinPeople[i];
                    break;
                }
            }
        }
        if (person) {
            state.selectedPerson = person;
            state.personDetail = null;
            $('#drPersonInfoBar').html('');
            hideModal('drWalkinModal');
            // Add to registrants list temporarily if not already there
            var exists = false;
            for (var j = 0; j < state.registrants.length; j++) {
                if (state.registrants[j].peopleId === pid) { exists = true; break; }
            }
            if (!exists) {
                state.registrants.unshift(person);
            }
            renderPeopleList($('#drSearchInput').val());
            renderAssignBar();
            if (state.selectedDestIdx === null) {
                showToast('Selected: ' + person.name + ' - Now tap a destination', 'info');
            }
            // Fetch full detail for walk-in too
            var sourceOrgId = (state.currentStation && state.currentStation.sourceOrgId) ? state.currentStation.sourceOrgId : 0;
            ajax('get_person_detail', {
                people_id: person.peopleId,
                source_org_id: sourceOrgId
            }, function(resp) {
                if (resp.success && state.selectedPerson && state.selectedPerson.peopleId === person.peopleId) {
                    state.personDetail = resp;
                    renderPersonInfoBar();
                }
            });
        }
    }

    function changeStation() {
        exitFullscreen();
        stopPolling();
        state.selectedPerson = null;
        state.selectedDestIdx = null;
        state.personDetail = null;
        state.registrants = [];
        state.destinations = [];
        state.assignedIds = {};
        state.mode = 'station_pick_scenario';
        render();
    }

    function showStationSwitcher() {
        var $dd = $('#drStationSwitcherDropdown');
        if ($dd.children().length) { hideStationSwitcher(); return; }
        var sc = state.currentScenario;
        if (!sc || !sc.stations || !sc.stations.length) return;
        var curId = state.currentStation ? state.currentStation.id : null;
        var html = '<div class="dr-station-switcher-backdrop" onclick="DRApp.hideStationSwitcher()"></div>';
        html += '<div class="dr-station-switcher">';
        html += '<div class="dr-station-switcher-title">Switch Station</div>';
        for (var i = 0; i < sc.stations.length; i++) {
            var s = sc.stations[i];
            var isActive = s.id === curId;
            html += '<button class="dr-station-switcher-item' + (isActive ? ' active' : '') + '" data-sid="' + escAttr(s.id) + '" onclick="' + (isActive ? 'DRApp.hideStationSwitcher()' : 'DRApp.quickSwitchStation(this.dataset.sid)') + '">';
            html += isActive ? '<i class="bi bi-check-lg"></i> ' : '';
            html += escHtml(s.name);
            html += '</button>';
        }
        html += '</div>';
        $dd.html(html);
        $('#drStationSwitchTrigger').addClass('open');
    }

    function hideStationSwitcher() {
        $('#drStationSwitcherDropdown').empty();
        $('#drStationSwitchTrigger').removeClass('open');
    }

    function quickSwitchStation(stationId) {
        hideStationSwitcher();
        var sc = state.currentScenario;
        if (!sc || !sc.stations) return;
        var newStation = null;
        for (var i = 0; i < sc.stations.length; i++) {
            if (sc.stations[i].id === stationId) { newStation = sc.stations[i]; break; }
        }
        if (!newStation) return;
        stopPolling();
        state.selectedPerson = null;
        state.selectedDestIdx = null;
        state.personDetail = null;
        state.registrants = [];
        state.destinations = [];
        state.assignedIds = {};
        state.searchTerm = '';
        state.currentStation = newStation;
        render();
    }

    // ---- Utilities ----
    function escHtml(s) {
        if (!s) return '';
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    function escAttr(s) {
        return escHtml(s).replace(/'/g, '&#39;');
    }

    // ---- Public API ----
    window.DRApp = {
        goLanding: function() { exitFullscreen(); stopPolling(); navigate('landing'); },
        goAdmin: function() { navigate('admin'); },
        goStation: function() { navigate('station_pick_scenario'); },
        goAdminList: function() { state.currentScenario = null; navigate('admin'); },
        createScenario: createScenario,
        editScenario: editScenario,
        deleteScenario: deleteScenario,
        saveScenario: saveScenario,
        addStation: addStation,
        editStation: editStation,
        removeStation: removeStation,
        saveStation: saveStation,
        addDest: addDest,
        addDestBulk: addDestBulk,
        searchDestOrgBulk: searchDestOrgBulk,
        toggleBulkPick: toggleBulkPick,
        removeBulkPick: removeBulkPick,
        updateBulkDefault: updateBulkDefault,
        cancelBulkAdd: cancelBulkAdd,
        commitBulkAdd: commitBulkAdd,
        editDest: editDest,
        removeDest: removeDest,
        saveDest: saveDest,
        cancelDestEdit: cancelDestEdit,
        clearStationSource: clearStationSource,
        updateStationName: updateStationName,
        clearDestOrg: clearDestOrg,
        searchSource: searchSource,
        searchDestOrg: searchDestOrg,
        selectOrg: selectOrg,
        selectOrgFromEl: selectOrgFromEl,
        pickScenario: pickScenario,
        pickStation: pickStation,
        startStation: startStation,
        navigate: navigate,
        selectPerson: selectPerson,
        selectDest: selectDest,
        changeDestSort: changeDestSort,
        assignToDest: assignToDest,
        filterPeople: filterPeople,
        clearSearch: clearSearch,
        switchPeopleTab: switchPeopleTab,
        dismissPersonInfoBar: dismissPersonInfoBar,
        toggleRegAnswers: toggleRegAnswers,
        showFindFriend: showFindFriend,
        doFindFriend: doFindFriend,
        ffProgramChanged: ffProgramChanged,
        searchAllPeople: searchAllPeople,
        doWalkinSearch: doWalkinSearch,
        selectWalkin: selectWalkin,
        changeStation: changeStation,
        showStationSwitcher: showStationSwitcher,
        hideStationSwitcher: hideStationSwitcher,
        quickSwitchStation: quickSwitchStation,
        toggleFullscreen: toggleFullscreen,
        showModal: showModal,
        hideModal: hideModal,
        loadRegQuestions: loadRegQuestions,
        togglePinQuestion: togglePinQuestion,
        updatePinLabel: updatePinLabel,
        toggleDestSubgroup: toggleDestSubgroup,
        toggleSubgroup: toggleSubgroup,
        applyAppUpdate: applyAppUpdate,
        toggleTestMode: toggleTestMode,
        clearAssignmentLog: clearAssignmentLog,
        resetTestAssignments: resetTestAssignments
    };

    // ---- Auto-update: version check, banner, apply ----
    function checkForAppUpdate() {
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', DC_API_BASE + '/script-versions', true);
            xhr.timeout = 5000;
            xhr.onreadystatechange = function() {
                if (xhr.readyState !== 4 || xhr.status !== 200) return;
                try {
                    var versions = JSON.parse(xhr.responseText);
                    var latest = versions[DC_SCRIPT_ID];
                    if (latest && latest !== APP_VERSION) {
                        APP_UPDATE_AVAILABLE = true;
                        APP_LATEST_VERSION = latest;
                        renderAppUpdateBanner();
                    }
                } catch(e) {}
            };
            xhr.send();
        } catch(e) {}
    }

    function renderAppUpdateBanner() {
        var b = document.getElementById('drAppUpdateBanner');
        if (!b || !APP_UPDATE_AVAILABLE) return;
        var h = '';
        h += '<div style="font-size:18px">&#128640;</div>';
        h += '<div style="flex:1;font-size:12px;color:#0078d4">';
        h += '<strong>Day of Registration update available</strong>';
        h += ' \\u2014 you have <code>v' + APP_VERSION + '</code>, latest is <code>v' + APP_LATEST_VERSION + '</code>. Your scenarios and logs are preserved.';
        h += '</div>';
        h += '<button id="drAppUpdateBtn" onclick="DRApp.applyAppUpdate()" style="white-space:nowrap;padding:6px 14px;background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer;">Update Now</button>';
        b.innerHTML = h;
        b.style.display = 'flex';
    }

    function applyAppUpdate() {
        if (!confirm('Update Day of Registration from v' + APP_VERSION + ' to v' + APP_LATEST_VERSION + '?\\n\\nYour scenarios and processing logs are stored separately and will be preserved.')) return;
        var btn = document.getElementById('drAppUpdateBtn');
        if (btn) { btn.disabled = true; btn.innerHTML = 'Updating...'; }
        $.post(scriptPath, {action: 'apply_update', script_name: SCRIPT_NAME}, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    alert(d.message || 'Updated! Reloading...');
                    window.location.reload(true);
                } else {
                    alert('Update failed: ' + (d.message || 'Unknown error'));
                    if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
                }
            } catch(e) {
                alert('Update failed (could not parse response)');
                if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
            }
        }).fail(function() {
            alert('Update failed (network error)');
            if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
        });
    }

    // ---- Init ----
    $(function() {
        render();
        checkForAppUpdate();
    });
})();
</script>

<div class="dr-root dr-container" style="padding-top:8px">
    <div id="drAppUpdateBanner" style="display:none;background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;align-items:center;gap:10px;"></div>
</div>

<div class="dr-root dr-container" id="drApp">
    <div class="dr-loading" style="display:block"><div class="dr-spinner"></div>Loading...</div>
</div>

<div class="dr-toast-container" id="drToastContainer"></div>
</body>
</html>'''

    print html
    model.Form = html
