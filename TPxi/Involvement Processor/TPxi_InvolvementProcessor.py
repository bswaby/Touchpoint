#Roles=Edit

"""
TPxi Involvement Processor
===========================
A comprehensive involvement management tool for processing registrants,
assigning people to target involvements, and tracking work across sessions.

Features:
- Search and filter involvements by name, ID, program, or division
- View registrants with registration question responses
- Advanced filtering by gender, age, subgroups, and question answers
- Bulk assignment of selected people to target involvements and subgroups
- Processed status tracking per registrant
- Effort management (named work sessions that persist state)
- Watchlist management for tracking specific people
- Save/load setup configurations for quick recall
- AJAX-based responsive UI with no page reloads

Written By: Ben Swaby
Email: bswaby@fbchtn.org
Version: 1.0
Date: January 2025

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python "TPxi_InvolvementProcessor" and paste all this code
4. Test and optionally add to menu via CustomReports
--Upload Instructions End--
"""

import json
import datetime

model.Header = 'Involvement Processor'

# ============================================================================
# AJAX HANDLER
# ============================================================================
# Handle POST requests (AJAX)
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -------------------------------------------------------------------------
    # Search Involvements
    # -------------------------------------------------------------------------
    if action == 'search_involvements':
        search_term = getattr(Data, 'search_term', '')
        division_id = getattr(Data, 'division_id', '')
        program_id = getattr(Data, 'program_id', '')

        where_clauses = ["o.OrganizationStatusId = 30"]  # Active only

        if search_term:
            # Escape single quotes
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

        try:
            results = q.QuerySql(sql)
            involvements = []
            for r in results:
                involvements.append({
                    'id': r.OrganizationId,
                    'name': r.OrganizationName,
                    'division': r.DivisionName or '',
                    'program': r.ProgramName or '',
                    'memberCount': r.MemberCount or 0
                })
            print json.dumps({'success': True, 'involvements': involvements})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Get Programs and Divisions for filter dropdowns
    # -------------------------------------------------------------------------
    elif action == 'get_filters':
        try:
            # Get programs
            prog_sql = """
                SELECT DISTINCT p.Id, p.Name
                FROM Program p
                JOIN Division d ON d.ProgId = p.Id
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY p.Name
            """
            programs = []
            for r in q.QuerySql(prog_sql):
                programs.append({'id': r.Id, 'name': r.Name})

            # Get divisions
            div_sql = """
                SELECT DISTINCT d.Id, d.Name, d.ProgId
                FROM Division d
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE o.OrganizationStatusId = 30
                ORDER BY d.Name
            """
            divisions = []
            for r in q.QuerySql(div_sql):
                divisions.append({'id': r.Id, 'name': r.Name, 'programId': r.ProgId})

            print json.dumps({'success': True, 'programs': programs, 'divisions': divisions})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Load Registrants with Registration Questions/Answers
    # Uses same approach as MissionsDashboard - handles both old and new registration systems
    # -------------------------------------------------------------------------
    elif action == 'load_registrants':
        import re
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)

                # Get organization members with birthdate for age-in-months calculation
                members_sql = """
                    SELECT
                        p.PeopleId,
                        p.Name2 as Name,
                        p.EmailAddress,
                        p.CellPhone,
                        p.Age,
                        p.BDate,
                        p.GenderId,
                        om.EnrollmentDate,
                        om.MemberTypeId,
                        mt.Description as MemberType
                    FROM OrganizationMembers om
                    JOIN People p ON om.PeopleId = p.PeopleId
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    WHERE om.OrganizationId = {0}
                    ORDER BY p.Name2
                """.format(org_id)

                members = []
                member_ids = []
                member_map = {}  # For quick lookup when adding subgroups
                for r in q.QuerySql(members_sql):
                    # Calculate age in months for young children
                    age_months = None
                    if r.BDate:
                        try:
                            today = datetime.date.today()
                            # Handle different date formats
                            if hasattr(r.BDate, 'year'):
                                # Already a datetime object
                                bdate = r.BDate if hasattr(r.BDate, 'month') else r.BDate.date()
                            else:
                                # String format - try multiple formats
                                bdate_str = str(r.BDate).strip()
                                bdate = None
                                # Try YYYY-MM-DD format first
                                try:
                                    bdate = datetime.datetime.strptime(bdate_str[:10], '%Y-%m-%d').date()
                                except:
                                    pass
                                # Try MM/DD/YYYY format
                                if not bdate:
                                    try:
                                        bdate = datetime.datetime.strptime(bdate_str[:10], '%m/%d/%Y').date()
                                    except:
                                        pass
                                # Try M/D/YYYY format (no leading zeros)
                                if not bdate:
                                    try:
                                        bdate = datetime.datetime.strptime(bdate_str.split()[0], '%m/%d/%Y').date()
                                    except:
                                        pass
                            if bdate:
                                age_months = (today.year - bdate.year) * 12 + (today.month - bdate.month)
                                if today.day < bdate.day:
                                    age_months -= 1
                        except:
                            # If date parsing fails completely, just skip age_months
                            age_months = None

                    member = {
                        'peopleId': r.PeopleId,
                        'name': r.Name,
                        'email': r.EmailAddress or '',
                        'phone': r.CellPhone or '',
                        'age': r.Age,
                        'ageMonths': age_months,
                        'gender': 'M' if r.GenderId == 1 else 'F' if r.GenderId == 2 else '',
                        'enrollmentDate': str(r.EnrollmentDate)[:10] if r.EnrollmentDate else '',
                        'memberType': r.MemberType or '',
                        'answers': {},
                        'subgroups': []  # Will be populated below
                    }
                    members.append(member)
                    member_ids.append(r.PeopleId)
                    member_map[r.PeopleId] = member

                # Get registration questions and answers using RegQuestion/RegAnswer tables (newer registrations)
                # This approach matches MissionsDashboard - gets questions through the answer relationships
                questions = []
                question_set = set()  # Track unique questions

                if member_ids:
                    # Query that gets both questions and answers in one go
                    answers_sql = """
                        SELECT
                            rp.PeopleId,
                            rq.RegQuestionId,
                            rq.Label as Question,
                            rq.[Order] as QuestionOrder,
                            ra.AnswerValue as Answer
                        FROM Registration r WITH (NOLOCK)
                        JOIN RegPeople rp WITH (NOLOCK) ON r.RegistrationId = rp.RegistrationId
                        LEFT JOIN RegAnswer ra WITH (NOLOCK) ON rp.RegPeopleId = ra.RegPeopleId
                        LEFT JOIN RegQuestion rq WITH (NOLOCK) ON ra.RegQuestionId = rq.RegQuestionId
                        WHERE r.OrganizationId = {0}
                          AND rp.PeopleId IN ({1})
                          AND rq.Label IS NOT NULL
                        ORDER BY rq.[Order]
                    """.format(org_id, ','.join(str(pid) for pid in member_ids))

                    try:
                        answers_result = q.QuerySql(answers_sql)
                        if answers_result:
                            for r in answers_result:
                                # Build questions list from actual answers found
                                if r.Question and r.Question not in question_set:
                                    questions.append({
                                        'id': str(r.RegQuestionId) if r.RegQuestionId else 'q_' + str(len(questions)),
                                        'text': r.Question
                                    })
                                    question_set.add(r.Question)

                                # Find the member and add their answer
                                for m in members:
                                    if m['peopleId'] == r.PeopleId:
                                        m['answers'][r.Question] = r.Answer or ''
                                        break
                    except:
                        pass

                # Also try RegistrationData table for older registrations (XML format)
                rd_sql = """
                    SELECT CAST(rd.Data AS NVARCHAR(MAX)) as XmlData
                    FROM RegistrationData rd WITH (NOLOCK)
                    WHERE rd.OrganizationId = {0}
                      AND rd.completed = 1
                    ORDER BY rd.Stamp DESC
                """.format(org_id)

                try:
                    rd_result = q.QuerySql(rd_sql)
                    if rd_result:
                        for row in rd_result:
                            if not row.XmlData:
                                continue
                            xml_data = row.XmlData

                            # Process each member
                            for m in members:
                                people_id = m['peopleId']
                                people_id_pattern = '<PeopleId>{0}</PeopleId>'.format(people_id)
                                if people_id_pattern not in xml_data:
                                    continue

                                # Find the person's section in XML
                                person_pattern = r'<OnlineRegPersonModel[^>]*>.*?<PeopleId>{0}</PeopleId>.*?</OnlineRegPersonModel>'.format(people_id)
                                person_match = re.search(person_pattern, xml_data, re.DOTALL)
                                if not person_match:
                                    continue

                                person_xml = person_match.group(0)

                                # Extract ExtraQuestion elements
                                extra_pattern = r'<ExtraQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</ExtraQuestion>'
                                for match in re.finditer(extra_pattern, person_xml):
                                    question = match.group(1)
                                    answer = match.group(2)
                                    if question and question not in m['answers']:
                                        m['answers'][question] = answer.strip() if answer else ''
                                        if question not in question_set:
                                            questions.append({'id': 'xml_' + question[:20], 'text': question})
                                            question_set.add(question)

                                # Extract Text elements
                                text_pattern = r'<Text[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</Text>'
                                for match in re.finditer(text_pattern, person_xml):
                                    question = match.group(1)
                                    answer = match.group(2)
                                    if question and question not in m['answers']:
                                        m['answers'][question] = answer.strip() if answer else ''
                                        if question not in question_set:
                                            questions.append({'id': 'xml_' + question[:20], 'text': question})
                                            question_set.add(question)

                                # Extract YesNoQuestion elements (convert True/False to Yes/No)
                                yn_pattern = r'<YesNoQuestion[^>]*\squestion="([^"]+)"[^>]*>([^<]*)</YesNoQuestion>'
                                for match in re.finditer(yn_pattern, person_xml):
                                    question = match.group(1)
                                    answer_val = match.group(2)
                                    if question and question not in m['answers']:
                                        # Convert True/False to Yes/No for readability
                                        answer = 'Yes' if answer_val.strip() == 'True' else 'No' if answer_val.strip() == 'False' else answer_val.strip()
                                        m['answers'][question] = answer
                                        if question not in question_set:
                                            questions.append({'id': 'xml_' + question[:20], 'text': question})
                                            question_set.add(question)
                except:
                    pass

                # Get subgroups for this org
                subgroups_sql = """
                    SELECT DISTINCT
                        sg.Name,
                        COUNT(DISTINCT ms.PeopleId) as MemberCount
                    FROM MemberTags sg
                    LEFT JOIN OrgMemMemTags ms ON sg.Id = ms.MemberTagId
                    WHERE sg.OrgId = {0}
                    GROUP BY sg.Id, sg.Name
                    ORDER BY sg.Name
                """.format(org_id)

                subgroups = []
                try:
                    sg_result = q.QuerySql(subgroups_sql)
                    if sg_result:
                        for r in sg_result:
                            subgroups.append({
                                'name': r.Name,
                                'count': r.MemberCount or 0
                            })
                except:
                    pass

                # Get subgroup membership for each member
                if member_ids:
                    member_subgroups_sql = """
                        SELECT
                            ommt.PeopleId,
                            mt.Name as SubgroupName
                        FROM OrgMemMemTags ommt
                        JOIN MemberTags mt ON ommt.MemberTagId = mt.Id
                        WHERE mt.OrgId = {0}
                          AND ommt.PeopleId IN ({1})
                    """.format(org_id, ','.join(str(pid) for pid in member_ids))

                    try:
                        msg_result = q.QuerySql(member_subgroups_sql)
                        if msg_result:
                            for r in msg_result:
                                if r.PeopleId in member_map:
                                    member_map[r.PeopleId]['subgroups'].append(r.SubgroupName)
                    except:
                        pass

                print json.dumps({
                    'success': True,
                    'members': members,
                    'questions': questions,
                    'subgroups': subgroups
                })
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Search Members within involvement (for match finding)
    # -------------------------------------------------------------------------
    elif action == 'search_members':
        org_id = getattr(Data, 'org_id', '')
        search_term = getattr(Data, 'search_term', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                where_clause = "om.OrganizationId = {0}".format(int(org_id))

                if search_term:
                    safe_term = search_term.replace("'", "''")
                    where_clause += """ AND (p.Name2 LIKE '%{0}%'
                        OR p.Name LIKE '%{0}%'
                        OR p.NickName LIKE '%{0}%')""".format(safe_term)

                sql = """
                    SELECT TOP 20
                        p.PeopleId,
                        p.Name2 as Name,
                        p.Age,
                        p.GenderId
                    FROM OrganizationMembers om
                    JOIN People p ON om.PeopleId = p.PeopleId
                    WHERE {0}
                    ORDER BY p.Name2
                """.format(where_clause)

                results = []
                for r in q.QuerySql(sql):
                    results.append({
                        'peopleId': r.PeopleId,
                        'name': r.Name,
                        'age': r.Age,
                        'gender': 'M' if r.GenderId == 1 else 'F' if r.GenderId == 2 else ''
                    })

                print json.dumps({'success': True, 'members': results})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Global People Search
    # -------------------------------------------------------------------------
    elif action == 'search_people':
        search_term = getattr(Data, 'search_term', '')

        if not search_term or len(search_term) < 2:
            print json.dumps({'success': False, 'message': 'Search term must be at least 2 characters'})
        else:
            try:
                safe_term = search_term.replace("'", "''")

                sql = """
                    SELECT TOP 30
                        p.PeopleId,
                        p.Name2 as Name,
                        p.EmailAddress,
                        p.Age,
                        p.GenderId,
                        ms.Description as MemberStatus
                    FROM People p
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    WHERE p.IsDeceased = 0
                    AND (p.Name2 LIKE '%{0}%'
                        OR p.Name LIKE '%{0}%'
                        OR p.NickName LIKE '%{0}%'
                        OR p.EmailAddress LIKE '%{0}%')
                    ORDER BY p.Name2
                """.format(safe_term)

                results = []
                for r in q.QuerySql(sql):
                    results.append({
                        'peopleId': r.PeopleId,
                        'name': r.Name,
                        'email': r.EmailAddress or '',
                        'age': r.Age,
                        'gender': 'M' if r.GenderId == 1 else 'F' if r.GenderId == 2 else '',
                        'status': r.MemberStatus or ''
                    })

                print json.dumps({'success': True, 'people': results})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Get Target Involvement Counts (for load balancing display)
    # -------------------------------------------------------------------------
    elif action == 'get_target_counts':
        org_ids = getattr(Data, 'org_ids', '')

        if not org_ids:
            print json.dumps({'success': False, 'message': 'Organization IDs required'})
        else:
            try:
                # Parse org IDs - convert to string first in case it comes in as int
                org_ids_str = str(org_ids)
                ids = [int(x.strip()) for x in org_ids_str.split(',') if x.strip()]

                results = []
                for org_id in ids:
                    try:
                        # Get org info and member count
                        org_sql = """
                            SELECT
                                o.OrganizationId,
                                o.OrganizationName,
                                (SELECT COUNT(*) FROM OrganizationMembers om
                                 WHERE om.OrganizationId = o.OrganizationId) as MemberCount
                            FROM Organizations o
                            WHERE o.OrganizationId = {0}
                        """.format(org_id)

                        org_result = q.QuerySql(org_sql)
                        org_list = list(org_result) if org_result else []

                        if len(org_list) > 0 and org_list[0]:
                            org = org_list[0]

                            # Get subgroup counts
                            sg_sql = """
                                SELECT
                                    sg.Name,
                                    COUNT(DISTINCT ms.PeopleId) as MemberCount
                                FROM MemberTags sg
                                LEFT JOIN OrgMemMemTags ms ON sg.Id = ms.MemberTagId
                                WHERE sg.OrgId = {0}
                                GROUP BY sg.Id, sg.Name
                                ORDER BY sg.Name
                            """.format(org_id)

                            subgroups = []
                            sg_result = q.QuerySql(sg_sql)
                            sg_list = list(sg_result) if sg_result else []
                            for sg in sg_list:
                                if sg and hasattr(sg, 'Name') and sg.Name:
                                    subgroups.append({
                                        'name': sg.Name,
                                        'count': sg.MemberCount or 0 if hasattr(sg, 'MemberCount') else 0
                                    })

                            results.append({
                                'orgId': org.OrganizationId if hasattr(org, 'OrganizationId') and org.OrganizationId else org_id,
                                'name': org.OrganizationName if hasattr(org, 'OrganizationName') and org.OrganizationName else 'Unknown',
                                'totalCount': org.MemberCount or 0 if hasattr(org, 'MemberCount') else 0,
                                'subgroups': subgroups
                            })
                    except Exception as inner_e:
                        # Log but continue with other orgs
                        results.append({
                            'orgId': org_id,
                            'name': 'Error loading org',
                            'totalCount': 0,
                            'subgroups': [],
                            'error': str(inner_e)
                        })

                print json.dumps({'success': True, 'targets': results, 'debug': {'org_ids_raw': str(org_ids), 'parsed_ids': ids, 'result_count': len(results)}})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Process Queue - Add people to target involvement/subgroups
    # -------------------------------------------------------------------------
    elif action == 'process_queue':
        people_ids = getattr(Data, 'people_ids', '')
        target_org_id = getattr(Data, 'target_org_id', '')
        subgroup_names = getattr(Data, 'subgroup_names', '')

        if not people_ids or not target_org_id:
            print json.dumps({'success': False, 'message': 'People and target organization required'})
        else:
            try:
                target_org_id = int(target_org_id)
                # Convert to string first in case they come in as different types
                ids = [int(x.strip()) for x in str(people_ids).split(',') if x.strip()]
                subgroups = [x.strip() for x in str(subgroup_names).split(',') if x.strip()] if subgroup_names else []

                processed = []
                errors = []

                for pid in ids:
                    try:
                        person = model.GetPerson(pid)
                        if person:
                            # Add to organization
                            if not model.InOrg(pid, target_org_id):
                                model.JoinOrg(target_org_id, person)

                            # Add to subgroups
                            for sg_name in subgroups:
                                model.AddSubGroup(pid, target_org_id, sg_name)

                            processed.append({
                                'peopleId': pid,
                                'name': person.Name2
                            })
                        else:
                            errors.append({'peopleId': pid, 'error': 'Person not found'})
                    except Exception as e:
                        errors.append({'peopleId': pid, 'error': str(e)})

                print json.dumps({
                    'success': True,
                    'processed': processed,
                    'errors': errors,
                    'message': 'Processed {0} people to org {1}, {2} subgroups assigned, {3} errors'.format(len(processed), target_org_id, len(subgroups), len(errors)),
                    'debug': {
                        'subgroups_requested': subgroups,
                        'target_org_id': target_org_id
                    }
                })
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Save Configuration
    # -------------------------------------------------------------------------
    elif action == 'save_config':
        org_id = getattr(Data, 'org_id', '')
        config_data = getattr(Data, 'config_data', '{}')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                # Store config in org extra value
                model.AddExtraValueTextOrg(org_id, "InvolvementProcessorConfig", config_data)
                print json.dumps({'success': True, 'message': 'Configuration saved'})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Load Configuration
    # -------------------------------------------------------------------------
    elif action == 'load_config':
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                config = model.ExtraValueTextOrg(org_id, "InvolvementProcessorConfig")
                if config:
                    print json.dumps({'success': True, 'config': json.loads(config)})
                else:
                    print json.dumps({'success': True, 'config': None})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Watchlist Management
    # -------------------------------------------------------------------------
    elif action == 'add_watchlist':
        org_id = getattr(Data, 'org_id', '')
        watcher_id = getattr(Data, 'watcher_id', '')
        requested_name = getattr(Data, 'requested_name', '')

        if not org_id or not watcher_id or not requested_name:
            print json.dumps({'success': False, 'message': 'All fields required'})
        else:
            try:
                org_id = int(org_id)

                # Get existing watchlist
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorWatchlist")
                watchlist = json.loads(existing) if existing else []

                # Add new entry
                watchlist.append({
                    'watcherId': int(watcher_id),
                    'requestedName': requested_name,
                    'addedDate': datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                    'status': 'pending'
                })

                # Save updated watchlist
                model.AddExtraValueTextOrg(org_id, "InvolvementProcessorWatchlist", json.dumps(watchlist))

                print json.dumps({'success': True, 'message': 'Added to watchlist', 'watchlist': watchlist})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    elif action == 'get_watchlist':
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorWatchlist")
                watchlist = json.loads(existing) if existing else []
                print json.dumps({'success': True, 'watchlist': watchlist})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    elif action == 'remove_watchlist':
        org_id = getattr(Data, 'org_id', '')
        index = getattr(Data, 'index', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                index = int(index)

                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorWatchlist")
                watchlist = json.loads(existing) if existing else []

                if 0 <= index < len(watchlist):
                    watchlist.pop(index)
                    model.AddExtraValueTextOrg(org_id, "InvolvementProcessorWatchlist", json.dumps(watchlist))

                print json.dumps({'success': True, 'watchlist': watchlist})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Effort Management - List all efforts for an organization
    # -------------------------------------------------------------------------
    elif action == 'list_efforts':
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorEfforts")
                # Debug: log what we found
                debug_info = {
                    'org_id': org_id,
                    'existing_raw': existing[:200] if existing else None,
                    'has_existing': bool(existing)
                }
                data = json.loads(existing) if existing else {'efforts': []}
                efforts = data.get('efforts', [])
                # Return just the summary info for listing
                effort_list = []
                for e in efforts:
                    effort_list.append({
                        'id': e.get('id'),
                        'name': e.get('name'),
                        'updatedAt': e.get('updatedAt'),
                        'processedCount': len(e.get('processed', []))
                    })
                print json.dumps({'success': True, 'efforts': effort_list, 'debug': debug_info})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Effort Management - Save an effort (create or update)
    # -------------------------------------------------------------------------
    elif action == 'save_effort':
        org_id = getattr(Data, 'org_id', '')
        effort_id = getattr(Data, 'effort_id', '')
        effort_name = getattr(Data, 'effort_name', '')
        effort_config = getattr(Data, 'effort_config', '{}')
        processed_json = getattr(Data, 'processed_json', '')
        processed_ids = getattr(Data, 'processed_ids', '')  # Legacy support

        if not org_id or not effort_name:
            print json.dumps({'success': False, 'message': 'Organization ID and effort name required'})
        else:
            try:
                org_id = int(org_id)
                config = json.loads(effort_config) if effort_config else {}

                # Handle new JSON format or legacy comma-separated IDs
                if processed_json:
                    processed = json.loads(processed_json) if processed_json else []
                elif processed_ids:
                    # Legacy format - convert to new format
                    processed = [{'peopleId': int(x.strip())} for x in str(processed_ids).split(',') if x.strip()]
                else:
                    processed = []

                # Get existing efforts
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorEfforts")
                data = json.loads(existing) if existing else {'efforts': []}
                efforts = data.get('efforts', [])

                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

                # Check if updating existing or creating new
                if effort_id:
                    # Update existing
                    found = False
                    for e in efforts:
                        if e.get('id') == effort_id:
                            e['name'] = effort_name
                            e['config'] = config
                            e['processed'] = processed
                            e['updatedAt'] = now
                            found = True
                            break
                    if not found:
                        print json.dumps({'success': False, 'message': 'Effort not found'})
                        # Early exit handled by the if/else structure
                else:
                    # Create new effort
                    import random
                    effort_id = 'effort_' + str(int(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))) + str(random.randint(100, 999))
                    efforts.append({
                        'id': effort_id,
                        'name': effort_name,
                        'createdAt': now,
                        'updatedAt': now,
                        'config': config,
                        'processed': processed
                    })

                # Save back
                data['efforts'] = efforts
                model.AddExtraValueTextOrg(org_id, "InvolvementProcessorEfforts", json.dumps(data))

                print json.dumps({'success': True, 'message': 'Effort saved', 'effortId': effort_id})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Effort Management - Load a specific effort
    # -------------------------------------------------------------------------
    elif action == 'load_effort':
        org_id = getattr(Data, 'org_id', '')
        effort_id = getattr(Data, 'effort_id', '')

        if not org_id or not effort_id:
            print json.dumps({'success': False, 'message': 'Organization ID and effort ID required'})
        else:
            try:
                org_id = int(org_id)
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorEfforts")
                data = json.loads(existing) if existing else {'efforts': []}
                efforts = data.get('efforts', [])

                effort = None
                for e in efforts:
                    if e.get('id') == effort_id:
                        effort = e
                        break

                if effort:
                    print json.dumps({'success': True, 'effort': effort})
                else:
                    print json.dumps({'success': False, 'message': 'Effort not found'})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Effort Management - Delete an effort
    # -------------------------------------------------------------------------
    elif action == 'delete_effort':
        org_id = getattr(Data, 'org_id', '')
        effort_id = getattr(Data, 'effort_id', '')

        if not org_id or not effort_id:
            print json.dumps({'success': False, 'message': 'Organization ID and effort ID required'})
        else:
            try:
                org_id = int(org_id)
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorEfforts")
                data = json.loads(existing) if existing else {'efforts': []}
                efforts = data.get('efforts', [])

                # Filter out the deleted effort
                new_efforts = [e for e in efforts if e.get('id') != effort_id]

                if len(new_efforts) < len(efforts):
                    data['efforts'] = new_efforts
                    model.AddExtraValueTextOrg(org_id, "InvolvementProcessorEfforts", json.dumps(data))
                    print json.dumps({'success': True, 'message': 'Effort deleted'})
                else:
                    print json.dumps({'success': False, 'message': 'Effort not found'})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    # -------------------------------------------------------------------------
    # Mark person as processed
    # -------------------------------------------------------------------------
    elif action == 'mark_processed':
        org_id = getattr(Data, 'org_id', '')
        people_id = getattr(Data, 'people_id', '')

        if not org_id or not people_id:
            print json.dumps({'success': False, 'message': 'Organization and People ID required'})
        else:
            try:
                org_id = int(org_id)
                people_id = int(people_id)

                # Get existing processed list
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorProcessed")
                processed = json.loads(existing) if existing else []

                if people_id not in processed:
                    processed.append(people_id)
                    model.AddExtraValueTextOrg(org_id, "InvolvementProcessorProcessed", json.dumps(processed))

                print json.dumps({'success': True, 'processed': processed})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    elif action == 'get_processed':
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                existing = model.ExtraValueTextOrg(org_id, "InvolvementProcessorProcessed")
                processed = json.loads(existing) if existing else []
                print json.dumps({'success': True, 'processed': processed})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    elif action == 'clear_processed':
        org_id = getattr(Data, 'org_id', '')

        if not org_id:
            print json.dumps({'success': False, 'message': 'Organization ID required'})
        else:
            try:
                org_id = int(org_id)
                model.AddExtraValueTextOrg(org_id, "InvolvementProcessorProcessed", "[]")
                print json.dumps({'success': True, 'message': 'Processed list cleared'})
            except Exception as e:
                print json.dumps({'success': False, 'message': str(e)})

    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# ============================================================================
# MAIN PAGE DISPLAY
# ============================================================================
else:
    html = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Auto-redirect to PyScriptForm (required for AJAX to work) -->
    <script>
    if (window.location.pathname.indexOf('/PyScript/') > -1) {
        window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
    }
    </script>
    <!-- Bootstrap Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
/* Involvement Processor - Scoped CSS (ip- prefix) */
.ip-root {
    --ip-primary: #3498db;
    --ip-secondary: #2ecc71;
    --ip-warning: #f39c12;
    --ip-danger: #e74c3c;
    --ip-dark: #2c3e50;
    --ip-light-bg: #f8f9fa;
}

.ip-main-container {
    max-width: 1600px;
    margin: 0 auto;
    padding: 20px;
}

.ip-panel {
    background: white;
    border-radius: 10px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.08);
    margin-bottom: 20px;
    overflow: hidden;
}

.ip-panel-header {
    background: linear-gradient(135deg, var(--ip-primary), #2980b9);
    color: white;
    padding: 15px 20px;
    font-weight: 600;
    font-size: 16px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.ip-panel-header .badge {
    background: rgba(255,255,255,0.2);
    font-size: 12px;
}

.ip-panel-body {
    padding: 20px;
}

.ip-setup-panel .ip-panel-header {
    background: linear-gradient(135deg, var(--ip-dark), #34495e);
}

.ip-queue-panel .ip-panel-header {
    background: linear-gradient(135deg, var(--ip-secondary), #27ae60);
}

.ip-target-panel .ip-panel-header {
    background: linear-gradient(135deg, var(--ip-warning), #e67e22);
}

.ip-watchlist-panel .ip-panel-header {
    background: linear-gradient(135deg, var(--ip-danger), #c0392b);
}

/* Selection Queue Sticky */
.ip-queue-sticky {
    position: sticky;
    top: 10px;
    z-index: 100;
}

/* Person Cards */
.ip-person-card {
    background: white;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 10px;
    transition: all 0.2s;
    cursor: pointer;
}

.ip-person-card:hover {
    border-color: var(--ip-primary);
    box-shadow: 0 2px 8px rgba(52, 152, 219, 0.2);
}

.ip-person-card.ip-selected {
    border-color: var(--ip-secondary);
    background: #e8f8f0;
}

.ip-person-card.ip-processed {
    opacity: 0.85;
    background: #e8f5e9;
    border-left: 3px solid #4caf50;
}

.ip-person-card .ip-assignment-info {
    margin-top: 6px;
    padding: 4px 8px;
    background: rgba(76, 175, 80, 0.1);
    border-radius: 4px;
    font-size: 12px;
    color: #2e7d32;
}

.ip-person-card .ip-assignment-info .text-success {
    color: #2e7d32 !important;
    font-weight: 500;
}

.ip-person-card .ip-assignment-info .text-info {
    color: #0277bd !important;
    font-weight: 500;
}

.ip-person-card .ip-name {
    font-weight: 600;
    color: var(--ip-dark);
    margin-bottom: 4px;
}

.ip-person-card .ip-details {
    font-size: 13px;
    color: #666;
}

.ip-person-card .ip-subgroups {
    margin-top: 4px;
    font-size: 11px;
    color: #6c757d;
}

.ip-person-card .ip-answers {
    margin-top: 8px;
    padding-top: 8px;
    border-top: 1px dashed #ddd;
}

.ip-person-card .ip-answer-item {
    font-size: 12px;
    color: #555;
    margin-bottom: 4px;
}

.ip-person-card .ip-answer-item strong {
    color: var(--ip-dark);
}

/* Queue Items */
.ip-queue-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 12px;
    background: #f8f9fa;
    border-radius: 6px;
    margin-bottom: 6px;
}

.ip-queue-item .ip-name {
    font-weight: 500;
}

.ip-queue-item .ip-remove-btn {
    color: var(--ip-danger);
    cursor: pointer;
    padding: 2px 6px;
}

.ip-queue-item .ip-remove-btn:hover {
    background: rgba(231, 76, 60, 0.1);
    border-radius: 4px;
}

/* Target Count Badges */
.ip-count-badge {
    display: inline-flex;
    align-items: center;
    background: #e9ecef;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 13px;
    margin: 3px;
}

.ip-count-badge.ip-highlight {
    background: var(--ip-secondary);
    color: white;
}

/* Search Box */
.ip-search-box {
    position: relative;
}

.ip-search-box input {
    padding-left: 40px;
}

.ip-search-box .ip-search-icon {
    position: absolute;
    left: 12px;
    top: 50%;
    transform: translateY(-50%);
    color: #999;
}

/* Tabs */
.ip-root .nav-tabs .nav-link {
    color: #666;
    border: none;
    padding: 10px 20px;
}

.ip-root .nav-tabs .nav-link.active {
    color: var(--ip-primary);
    border-bottom: 2px solid var(--ip-primary);
    background: transparent;
}

/* Loading Spinner */
.ip-loading-overlay {
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(255,255,255,0.8);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 10;
}

/* Involvement Selector */
.ip-involvement-item {
    padding: 10px 15px;
    border-bottom: 1px solid #eee;
    cursor: pointer;
    transition: background 0.15s;
}

.ip-involvement-item:hover {
    background: #f5f5f5;
}

.ip-involvement-item.ip-selected {
    background: #e3f2fd;
    border-left: 3px solid var(--ip-primary);
}

.ip-involvement-item .ip-org-name {
    font-weight: 600;
    color: var(--ip-dark);
}

.ip-involvement-item .ip-org-meta {
    font-size: 12px;
    color: #888;
}

/* Question Toggles */
.ip-question-toggle {
    display: flex;
    align-items: center;
    padding: 8px;
    background: #f8f9fa;
    border-radius: 6px;
    margin-bottom: 5px;
}

.ip-question-toggle input {
    margin-right: 10px;
}

.ip-question-toggle label {
    font-size: 13px;
    margin-bottom: 0;
    cursor: pointer;
}

/* Action Buttons */
.ip-action-btn {
    padding: 12px 24px;
    border-radius: 8px;
    font-weight: 600;
    transition: all 0.2s;
}

.ip-action-btn:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

/* Toast Notifications */
.ip-toast-container {
    position: fixed;
    top: 20px;
    right: 20px;
    z-index: 9999;
}

/* Responsive Grid */
.ip-registrant-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
    gap: 15px;
}

/* Empty State */
.ip-empty-state {
    text-align: center;
    padding: 40px;
    color: #999;
}

.ip-empty-state i {
    font-size: 48px;
    margin-bottom: 15px;
}

/* Filter Chips */
.ip-filter-chip {
    display: inline-flex;
    align-items: center;
    background: #e9ecef;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 13px;
    margin: 2px;
}

.ip-filter-chip .ip-remove {
    margin-left: 8px;
    cursor: pointer;
    color: #999;
}

.ip-filter-chip .ip-remove:hover {
    color: var(--ip-danger);
}

/* ============================================================================
   Bootstrap-like Components Scoped to .ip-root
   ============================================================================ */

/* Grid System */
.ip-root .row { display: flex; flex-wrap: wrap; margin: 0 -12px; }
.ip-root .col-lg-8 { flex: 0 0 66.66667%; max-width: 66.66667%; padding: 0 12px; }
.ip-root .col-lg-4 { flex: 0 0 33.33333%; max-width: 33.33333%; padding: 0 12px; }
@media (max-width: 991.98px) {
    .ip-root .col-lg-8, .ip-root .col-lg-4 { flex: 0 0 100%; max-width: 100%; }
}

/* Buttons */
.ip-root .btn {
    display: inline-block;
    font-weight: 400;
    text-align: center;
    vertical-align: middle;
    cursor: pointer;
    user-select: none;
    padding: 6px 12px;
    font-size: 14px;
    line-height: 1.5;
    border-radius: 4px;
    border: 1px solid transparent;
    transition: all 0.15s ease-in-out;
}
.ip-root .btn:hover { opacity: 0.85; }
.ip-root .btn:disabled { opacity: 0.65; cursor: not-allowed; }
.ip-root .btn-primary { background: var(--ip-primary); color: white; border-color: var(--ip-primary); }
.ip-root .btn-secondary { background: #6c757d; color: white; border-color: #6c757d; }
.ip-root .btn-success { background: var(--ip-secondary); color: white; border-color: var(--ip-secondary); }
.ip-root .btn-danger { background: var(--ip-danger); color: white; border-color: var(--ip-danger); }
.ip-root .btn-warning { background: var(--ip-warning); color: white; border-color: var(--ip-warning); }
.ip-root .btn-light { background: #f8f9fa; color: #212529; border-color: #f8f9fa; }
.ip-root .btn-outline-secondary { background: transparent; color: #6c757d; border-color: #6c757d; }
.ip-root .btn-outline-secondary:hover { background: #6c757d; color: white; }
.ip-root .btn-outline-danger { background: transparent; color: var(--ip-danger); border-color: var(--ip-danger); }
.ip-root .btn-outline-danger:hover { background: var(--ip-danger); color: white; }
.ip-root .btn-outline-light { background: transparent; color: #f8f9fa; border-color: rgba(255,255,255,0.5); }
.ip-root .btn-outline-light:hover { background: rgba(255,255,255,0.1); }
.ip-root .btn-sm { padding: 4px 8px; font-size: 12px; }
.ip-root .btn-close { box-sizing: content-box; width: 1em; height: 1em; padding: 0.25em; background: transparent url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3e%3cpath d='M.293.293a1 1 0 011.414 0L8 6.586 14.293.293a1 1 0 111.414 1.414L9.414 8l6.293 6.293a1 1 0 01-1.414 1.414L8 9.414l-6.293 6.293a1 1 0 01-1.414-1.414L6.586 8 .293 1.707a1 1 0 010-1.414z'/%3e%3c/svg%3e") center/1em no-repeat; border: 0; border-radius: 4px; opacity: 0.5; cursor: pointer; }
.ip-root .btn-close:hover { opacity: 1; }
.ip-root .btn-close-white { filter: invert(1) grayscale(100%) brightness(200%); }
.ip-root .btn-group { display: inline-flex; }
.ip-root .btn-group .btn { border-radius: 0; }
.ip-root .btn-group .btn:first-child { border-radius: 4px 0 0 4px; }
.ip-root .btn-group .btn:last-child { border-radius: 0 4px 4px 0; }

/* Forms */
.ip-root .form-control, .ip-root .form-select {
    display: block;
    width: 100%;
    padding: 6px 12px;
    font-size: 14px;
    line-height: 1.5;
    color: #212529;
    background-color: #fff;
    background-clip: padding-box;
    border: 1px solid #ced4da;
    border-radius: 4px;
    transition: border-color 0.15s ease-in-out;
}
.ip-root .form-control:focus, .ip-root .form-select:focus {
    border-color: var(--ip-primary);
    outline: 0;
    box-shadow: 0 0 0 2px rgba(52, 152, 219, 0.25);
}
.ip-root .form-control-sm, .ip-root .form-select-sm { padding: 4px 8px; font-size: 12px; }
.ip-root .form-label { display: block; margin-bottom: 4px; font-weight: 500; }
.ip-root .form-check { display: flex; align-items: center; padding-left: 0; margin-bottom: 8px; }
.ip-root .form-check-input { width: 16px; height: 16px; margin-right: 8px; cursor: pointer; }
.ip-root .form-check-label { cursor: pointer; }
.ip-root .input-group { display: flex; }
.ip-root .input-group .form-control { border-radius: 4px 0 0 4px; }
.ip-root .input-group .btn { border-radius: 0 4px 4px 0; }

/* Badges */
.ip-root .badge {
    display: inline-block;
    padding: 4px 8px;
    font-size: 11px;
    font-weight: 600;
    line-height: 1;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    border-radius: 4px;
}
.ip-root .bg-secondary { background: #6c757d; color: white; }
.ip-root .bg-success { background: var(--ip-secondary); color: white; }
.ip-root .bg-danger { background: var(--ip-danger); color: white; }
.ip-root .bg-warning { background: var(--ip-warning); color: white; }
.ip-root .bg-info { background: #17a2b8; color: white; }
.ip-root .bg-light { background: #f8f9fa; }
.ip-root .text-dark { color: #212529 !important; }

/* Nav Tabs */
.ip-root .nav { display: flex; flex-wrap: wrap; padding-left: 0; margin-bottom: 0; list-style: none; }
.ip-root .nav-tabs { border-bottom: 1px solid #dee2e6; }
.ip-root .nav-link { display: block; padding: 8px 16px; text-decoration: none; color: #6c757d; cursor: pointer; }
.ip-root .nav-link.active { color: var(--ip-primary); border-bottom: 2px solid var(--ip-primary); }
.ip-root .tab-content > .tab-pane { display: none; }
.ip-root .tab-content > .active { display: block; }

/* Spinners */
.ip-root .spinner-border {
    display: inline-block;
    width: 2rem;
    height: 2rem;
    vertical-align: text-bottom;
    border: 0.25em solid currentColor;
    border-right-color: transparent;
    border-radius: 50%;
    animation: ip-spinner 0.75s linear infinite;
}
.ip-root .spinner-border-sm { width: 1rem; height: 1rem; border-width: 0.2em; }
.ip-root .text-primary { color: var(--ip-primary) !important; }
@keyframes ip-spinner { to { transform: rotate(360deg); } }

/* Modals */
.ip-root .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: 1050; overflow-x: hidden; overflow-y: auto; outline: 0; background: rgba(0,0,0,0.5); opacity: 0; transition: opacity 0.15s linear; }
.ip-root .modal.show { display: block; opacity: 1; }
.ip-root .modal.fade .modal-dialog { transform: translate(0, -50px); transition: transform 0.3s ease-out; }
.ip-root .modal.fade.show .modal-dialog { transform: none; }
.ip-root .modal-dialog { position: relative; width: auto; max-width: 500px; margin: 1.75rem auto; pointer-events: none; }
.ip-root .modal-dialog.modal-lg { max-width: 800px; }
.ip-root .modal-content { position: relative; display: flex; flex-direction: column; width: 100%; pointer-events: auto; background-color: #fff; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
.ip-root .modal-header { display: flex; align-items: center; justify-content: space-between; padding: 16px 20px; border-bottom: 1px solid #dee2e6; }
.ip-root .modal-title { margin: 0; font-size: 18px; font-weight: 600; }
.ip-root .modal-body { position: relative; flex: 1 1 auto; padding: 20px; }
.ip-root .modal-footer { display: flex; align-items: center; justify-content: flex-end; padding: 12px 20px; border-top: 1px solid #dee2e6; gap: 8px; }

/* Toast */
.ip-root .toast { position: relative; width: 350px; max-width: 100%; font-size: 14px; background-color: #fff; background-clip: padding-box; border: 1px solid rgba(0,0,0,0.1); box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-radius: 6px; margin-bottom: 8px; }
.ip-root .toast .d-flex { display: flex; align-items: center; }
.ip-root .toast-body { padding: 12px 16px; flex: 1; }

/* Utility Classes */
.ip-root .d-flex { display: flex !important; }
.ip-root .d-block { display: block !important; }
.ip-root .d-none { display: none !important; }
.ip-root .justify-content-between { justify-content: space-between !important; }
.ip-root .align-items-center { align-items: center !important; }
.ip-root .text-center { text-align: center !important; }
.ip-root .text-muted { color: #6c757d !important; }
.ip-root .text-success { color: var(--ip-secondary) !important; }
.ip-root .text-white { color: #fff !important; }
.ip-root .fw-bold { font-weight: 700 !important; }
.ip-root .small { font-size: 0.875em !important; }
.ip-root .w-100 { width: 100% !important; }
.ip-root .float-end { float: right !important; }
.ip-root .mb-0 { margin-bottom: 0 !important; }
.ip-root .mb-1 { margin-bottom: 0.25rem !important; }
.ip-root .mb-2 { margin-bottom: 0.5rem !important; }
.ip-root .mb-3 { margin-bottom: 1rem !important; }
.ip-root .mb-4 { margin-bottom: 1.5rem !important; }
.ip-root .mt-1 { margin-top: 0.25rem !important; }
.ip-root .mt-2 { margin-top: 0.5rem !important; }
.ip-root .mt-3 { margin-top: 1rem !important; }
.ip-root .me-1 { margin-right: 0.25rem !important; }
.ip-root .me-2 { margin-right: 0.5rem !important; }
.ip-root .ms-2 { margin-left: 0.5rem !important; }
.ip-root .py-3 { padding-top: 1rem !important; padding-bottom: 1rem !important; }
.ip-root .py-4 { padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }
.ip-root .p-2 { padding: 0.5rem !important; }
.ip-root .p-3 { padding: 1rem !important; }
    </style>
</head>
<body>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
// Catch any JS errors and display them visibly
window.onerror = function(msg, url, lineNo, columnNo, error) {
    console.error('JS Error:', msg, 'at line', lineNo, 'col', columnNo);
    // Show error visibly on page
    var errDiv = document.getElementById('jsDebugOutput');
    if (errDiv) {
        errDiv.innerHTML += '<div class="text-danger">JS ERROR: ' + msg + ' at line ' + lineNo + '</div>';
        errDiv.style.display = 'block';
    }
    return false;
};
console.log('=== INVOLVEMENT PROCESSOR JS STARTING ===');
console.log('Current URL:', window.location.href);
console.log('Pathname:', window.location.pathname);

// Global State
var state = {
    sourceOrgId: null,
    sourceOrgName: '',
    registrants: [],
    questions: [],
    subgroups: [],
    selectedQuestions: [],
    queue: [],
    processed: [],
    targetOrgId: null,
    targetOrgName: '',
    targetSubgroups: [],
    selectedSubgroups: [],
    programs: [],
    divisions: [],
    watchlist: [],
    // Effort tracking
    currentEffort: null,  // { name: 'Cabin Assignments', id: 'effort_123' }
    efforts: [],          // List of available efforts for current source
    effortDirty: false    // Track if current effort has unsaved changes
};

var searchTimeout = null;
var scriptUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
console.log('scriptUrl set to:', scriptUrl);

// ============================================================================
// Processed Person Helpers
// ============================================================================
// state.processed is now an array of objects:
// [{ peopleId: 123, targetOrgId: 456, targetOrgName: "Camp", subgroups: ["Cabin A"], processedAt: "2026-01-20 20:46" }]

function isProcessed(peopleId) {
    return state.processed.some(function(p) {
        // Handle both old format (just peopleId number) and new format (object)
        return (typeof p === 'object' ? p.peopleId : p) === peopleId;
    });
}

function getProcessedInfo(peopleId) {
    for (var i = 0; i < state.processed.length; i++) {
        var p = state.processed[i];
        if (typeof p === 'object' && p.peopleId === peopleId) {
            return p;
        } else if (p === peopleId) {
            // Old format - just return minimal info
            return { peopleId: peopleId };
        }
    }
    return null;
}

function addProcessedPerson(peopleId, targetOrgId, targetOrgName, subgroups) {
    // Check if already processed
    if (isProcessed(peopleId)) return;

    var now = new Date();
    var dateStr = now.getFullYear() + '-' +
        String(now.getMonth() + 1).padStart(2, '0') + '-' +
        String(now.getDate()).padStart(2, '0') + ' ' +
        String(now.getHours()).padStart(2, '0') + ':' +
        String(now.getMinutes()).padStart(2, '0');

    state.processed.push({
        peopleId: peopleId,
        targetOrgId: targetOrgId,
        targetOrgName: targetOrgName,
        subgroups: subgroups || [],
        processedAt: dateStr
    });
}

function getProcessedIds() {
    // Convert processed array to comma-separated IDs for saving
    return state.processed.map(function(p) {
        return typeof p === 'object' ? p.peopleId : p;
    }).join(',');
}

function getProcessedJson() {
    // Get full processed array as JSON for saving
    return JSON.stringify(state.processed);
}
console.log('jQuery available:', typeof $ !== 'undefined');
console.log('jQuery version:', typeof $ !== 'undefined' ? $.fn.jquery : 'NOT LOADED');

// ============================================================================
// AJAX Helper
// ============================================================================
function ajax(action, params, callback) {
    console.log('ajax() called with action:', action);
    params = params || {};
    params.ajax = 'true';
    params.action = action;

    console.log('AJAX URL:', scriptUrl);
    console.log('AJAX params:', params);

    $.ajax({
        url: scriptUrl,
        type: 'POST',
        data: params,
        success: function(response) {
            console.log('AJAX success - raw response:', response);
            console.log('Response length:', response ? response.length : 0);
            try {
                var data = JSON.parse(response);
                console.log('Parsed response data:', data);
                if (callback) callback(data);
            } catch(e) {
                console.error('Parse error:', e, response);
                showToast('Error processing response', 'danger');
            }
        },
        error: function(xhr, status, error) {
            console.error('AJAX error:', error);
            console.error('XHR status:', status);
            console.error('XHR responseText:', xhr.responseText);
            showToast('Network error: ' + error, 'danger');
        }
    });
}

// ============================================================================
// Utility Functions
// ============================================================================
function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// ============================================================================
// Toast Notifications
// ============================================================================
function showToast(message, type) {
    type = type || 'info';
    var bgClass = {
        'success': 'bg-success',
        'danger': 'bg-danger',
        'warning': 'bg-warning',
        'info': 'bg-info'
    }[type] || 'bg-info';

    var toast = $('<div class="toast align-items-center text-white ' + bgClass + ' border-0" role="alert">' +
        '<div class="d-flex">' +
        '<div class="toast-body">' + message + '</div>' +
        '<button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>' +
        '</div></div>');

    $('#toastContainer').append(toast);
    var bsToast = new bootstrap.Toast(toast[0], { delay: 3000 });
    bsToast.show();
    toast.on('hidden.bs.toast', function() { $(this).remove(); });
}

// ============================================================================
// Initialization
// ============================================================================
$(document).ready(function() {
    console.log('Document ready - binding event handlers');
    console.log('involvementSearch element found:', $('#involvementSearch').length > 0);

    loadFilters();

    // Bind event handlers using jQuery (more reliable than inline handlers)
    $('#involvementSearch').on('keyup', debounceSearch);
    console.log('Event handlers bound');
    $('#programFilter').on('change', filterInvolvements);
    $('#divisionFilter').on('change', filterInvolvements);
    $('#registrantSearch').on('keyup', filterRegistrantList);
    $('#memberSearch').on('keyup', function() { searchMembers(); });
    $('#globalSearch').on('keyup', function() { searchGlobalPeople(); });
    $('#targetOrgSearch').on('keyup', function() { searchTargetOrg(); });

    // Radio button handlers
    $('input[name="targetType"]').on('change', updateTargetMode);

    // Event delegation for dynamically created elements
    $(document).on('change', '[data-qid]', function() {
        toggleQuestion($(this).data('qid'));
    });
    $(document).on('change', '[data-sgname]', function() {
        toggleSubgroup($(this).data('sgname'));
    });
});

function loadFilters() {
    ajax('get_filters', {}, function(response) {
        if (response.success) {
            state.programs = response.programs;
            state.divisions = response.divisions;

            var progHtml = '<option value="">All Programs</option>';
            response.programs.forEach(function(p) {
                progHtml += '<option value="' + p.id + '">' + p.name + '</option>';
            });
            $('#programFilter').html(progHtml);

            var divHtml = '<option value="">All Divisions</option>';
            response.divisions.forEach(function(d) {
                divHtml += '<option value="' + d.id + '">' + d.name + '</option>';
            });
            $('#divisionFilter').html(divHtml);
        }
    });
}

// ============================================================================
// Effort Management
// ============================================================================
function loadEfforts(orgId, skipPrompt) {
    console.log('loadEfforts called with orgId:', orgId, 'skipPrompt:', skipPrompt);
    if (!orgId) {
        console.error('loadEfforts called without orgId');
        return;
    }

    ajax('list_efforts', { org_id: orgId }, function(response) {
        console.log('list_efforts response:', response);
        console.log('efforts count:', response.efforts ? response.efforts.length : 0);
        if (response.success) {
            state.efforts = response.efforts || [];
            console.log('state.efforts:', state.efforts);

            // If there are existing efforts and we're not skipping the prompt, show the choice modal
            if (state.efforts.length > 0 && !skipPrompt) {
                console.log('Should show continue modal - efforts found:', state.efforts.length);
                showContinueEffortModal();
            } else if (state.efforts.length === 0) {
                console.log('No efforts found, auto-creating default');
                // No previous efforts - auto-create a default one
                autoCreateDefaultEffort();
            }

            $('#effortPanel').show();
            updateEffortDisplay();
        } else {
            console.error('Failed to load efforts:', response.message);
            state.efforts = [];
            autoCreateDefaultEffort();
            $('#effortPanel').show();
        }
    });
}

function showContinueEffortModal() {
    console.log('showContinueEffortModal called, efforts:', state.efforts);
    try {
        // Build list of previous efforts
        var html = '';
        state.efforts.forEach(function(e) {
            console.log('Building effort item:', e);
            var dateStr = e.updatedAt || 'Unknown date';
            html += '<a href="#" class="list-group-item list-group-item-action" data-effort-id="' + e.id + '">';
            html += '<div class="d-flex justify-content-between align-items-center">';
            html += '<div>';
            html += '<strong>' + escapeHtml(e.name) + '</strong>';
            html += '<br><small class="text-muted">' + (e.processedCount || 0) + ' processed - Last updated: ' + dateStr + '</small>';
            html += '</div>';
            html += '<i class="bi bi-chevron-right"></i>';
            html += '</div>';
            html += '</a>';
        });
        console.log('Modal HTML built, length:', html.length);
        $('#previousEffortsList').html(html);

        // Bind click handlers using jQuery (avoids quote escaping issues)
        $('#previousEffortsList a').on('click', function(ev) {
            ev.preventDefault();
            var effortId = $(this).data('effort-id');
            console.log('Effort clicked:', effortId);
            selectEffortFromModal(effortId);
        });

        var modalEl = document.getElementById('continueEffortModal');
        console.log('Modal element:', modalEl);
        if (modalEl) {
            // Check if there's already an instance
            var existingModal = bootstrap.Modal.getInstance(modalEl);
            console.log('Existing modal instance:', existingModal);

            if (existingModal) {
                console.log('Using existing modal instance');
                existingModal.show();
            } else {
                console.log('Creating new modal instance');
                var modal = new bootstrap.Modal(modalEl);
                modal.show();
            }
            console.log('Modal show() called');

            // Debug: Check modal state after a brief delay
            setTimeout(function() {
                console.log('Modal classList after show:', modalEl.className);
                console.log('Modal display style:', window.getComputedStyle(modalEl).display);
            }, 500);
        } else {
            console.error('continueEffortModal element not found');
        }
    } catch (err) {
        console.error('Error in showContinueEffortModal:', err);
    }
}

function selectEffortFromModal(effortId) {
    console.log('selectEffortFromModal called with:', effortId);
    try {
        // Close modal
        var modalEl = document.getElementById('continueEffortModal');
        var modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (modalInstance) modalInstance.hide();

        // Load the selected effort
        loadEffort(effortId);
    } catch (err) {
        console.error('Error in selectEffortFromModal:', err);
    }
}

function startFreshEffort() {
    console.log('startFreshEffort called');
    try {
        // Close modal
        var modalEl = document.getElementById('continueEffortModal');
        var modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (modalInstance) modalInstance.hide();

        // Show the name modal for the new effort
        var today = new Date();
        var defaultName = 'Session ' + (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();
        $('#newEffortName').val(defaultName);
        var newModalEl = document.getElementById('newEffortModal');
        if (newModalEl) {
            var modal = new bootstrap.Modal(newModalEl);
            modal.show();
        }
    } catch (err) {
        console.error('Error in startFreshEffort:', err);
    }
}

function autoCreateDefaultEffort() {
    console.log('autoCreateDefaultEffort called, sourceOrgId:', state.sourceOrgId);
    if (!state.sourceOrgId) {
        console.error('autoCreateDefaultEffort called without sourceOrgId');
        return;
    }

    // Create a default effort automatically
    var today = new Date();
    var defaultName = 'Session ' + (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();

    var config = {
        selectedQuestions: state.selectedQuestions,
        selectedSubgroups: state.selectedSubgroups
    };

    ajax('save_effort', {
        org_id: state.sourceOrgId,
        effort_id: '',
        effort_name: defaultName,
        effort_config: JSON.stringify(config),
        processed_json: '[]'
    }, function(response) {
        console.log('save_effort response:', response);
        if (response.success) {
            state.currentEffort = {
                id: response.effortId,
                name: defaultName
            };
            state.effortDirty = false;
            updateEffortDisplay();
        }
    });
}

function loadEffort(effortId) {
    ajax('load_effort', { org_id: state.sourceOrgId, effort_id: effortId }, function(response) {
        if (response.success && response.effort) {
            var effort = response.effort;
            state.currentEffort = {
                id: effort.id,
                name: effort.name
            };
            state.processed = effort.processed || [];
            state.effortDirty = false;

            // Apply config if available
            if (effort.config) {
                if (effort.config.selectedQuestions) {
                    state.selectedQuestions = effort.config.selectedQuestions;
                    renderQuestionToggles();
                }
                if (effort.config.selectedSubgroups) {
                    state.selectedSubgroups = effort.config.selectedSubgroups;
                }
            }

            updateEffortDisplay();
            renderRegistrantList();
            showToast('Loaded: ' + effort.name, 'success');
        } else {
            showToast('Failed to load effort', 'danger');
        }
    });
}

function showEffortPicker() {
    // Refresh efforts and show picker modal
    ajax('list_efforts', { org_id: state.sourceOrgId }, function(response) {
        if (response.success) {
            state.efforts = response.efforts || [];
            showContinueEffortModal();
        }
    });
}

function renameCurrentEffort() {
    if (!state.currentEffort) {
        showToast('No effort to rename', 'warning');
        return;
    }

    var newName = prompt('Enter new name for this effort:', state.currentEffort.name);
    if (newName && newName.trim() && newName.trim() !== state.currentEffort.name) {
        state.currentEffort.name = newName.trim();
        state.effortDirty = true;
        autoSaveEffort();
        updateEffortDisplay();
    }
}

function saveCurrentEffort() {
    if (!state.sourceOrgId || !state.currentEffort) {
        return;
    }

    var config = {
        selectedQuestions: state.selectedQuestions,
        selectedSubgroups: state.selectedSubgroups
    };

    ajax('save_effort', {
        org_id: state.sourceOrgId,
        effort_id: state.currentEffort.id,
        effort_name: state.currentEffort.name,
        effort_config: JSON.stringify(config),
        processed_json: getProcessedJson()
    }, function(response) {
        if (response.success) {
            state.effortDirty = false;
            updateEffortDisplay();
        }
    });
}

function showNewEffortModal() {
    var today = new Date();
    var defaultName = 'Session ' + (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();
    $('#newEffortName').val(defaultName);
    var modal = new bootstrap.Modal(document.getElementById('newEffortModal'));
    modal.show();
}

function createNewEffort() {
    var name = $('#newEffortName').val().trim();
    if (!name) {
        showToast('Please enter an effort name', 'warning');
        return;
    }

    if (!state.sourceOrgId) {
        showToast('Select a source involvement first', 'warning');
        return;
    }

    var config = {
        selectedQuestions: state.selectedQuestions,
        selectedSubgroups: state.selectedSubgroups
    };

    ajax('save_effort', {
        org_id: state.sourceOrgId,
        effort_id: '',
        effort_name: name,
        effort_config: JSON.stringify(config),
        processed_ids: ''  // Start fresh with no processed
    }, function(response) {
        if (response.success) {
            // Close modal
            bootstrap.Modal.getInstance(document.getElementById('newEffortModal')).hide();

            // Set as current effort with fresh processed list
            state.currentEffort = {
                id: response.effortId,
                name: name
            };
            state.processed = [];
            state.effortDirty = false;

            // Refresh efforts list
            ajax('list_efforts', { org_id: state.sourceOrgId }, function(resp) {
                if (resp.success) state.efforts = resp.efforts || [];
            });

            updateEffortDisplay();
            renderRegistrantList();
            showToast('Created: ' + name, 'success');
        } else {
            showToast('Failed to create: ' + response.message, 'danger');
        }
    });
}

function updateEffortDisplay() {
    // Show/hide dirty indicator
    if (state.effortDirty || state.queue.length > 0) {
        $('#effortDirtyIndicator').show();
    } else {
        $('#effortDirtyIndicator').hide();
    }

    if (state.currentEffort) {
        $('#effortName').text(state.currentEffort.name);
        var pending = state.registrants.filter(function(r) {
            return !isProcessed(r.peopleId);
        }).length;
        var total = state.registrants.length;
        var processed = state.processed.length;
        $('#effortProgress').text(pending + ' pending, ' + processed + ' processed of ' + total + ' total');
    } else {
        $('#effortName').text('No effort selected');
        $('#effortProgress').text('');
    }
}

function clearEffortState() {
    state.currentEffort = null;
    state.efforts = [];
    state.effortDirty = false;
    state.processed = [];
    $('#effortPanel').hide();
    $('#effortDirtyIndicator').hide();
}

// ============================================================================
// Involvement Search
// ============================================================================
function debounceSearch() {
    console.log('debounceSearch called');
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(searchInvolvements, 300);
}

function searchInvolvements() {
    console.log('searchInvolvements called');
    var term = $('#involvementSearch').val();
    var progId = $('#programFilter').val();
    var divId = $('#divisionFilter').val();
    console.log('Search params:', { term: term, progId: progId, divId: divId });

    if (!term && !progId && !divId) {
        $('#involvementList').html('<div class="ip-empty-state py-4"><i class="bi bi-search d-block"></i><span>Search for an involvement</span></div>');
        return;
    }

    $('#involvementList').html('<div class="text-center py-3"><div class="spinner-border spinner-border-sm"></div></div>');

    console.log('Making AJAX request for search_involvements');
    ajax('search_involvements', {
        search_term: term,
        program_id: progId,
        division_id: divId
    }, function(response) {
        if (response.success) {
            renderInvolvementList(response.involvements);
        } else {
            showToast(response.message, 'danger');
        }
    });
}

function filterInvolvements() {
    // Update divisions based on program selection
    var progId = $('#programFilter').val();
    if (progId) {
        var divHtml = '<option value="">All Divisions</option>';
        state.divisions.forEach(function(d) {
            if (d.programId == progId) {
                divHtml += '<option value="' + d.id + '">' + d.name + '</option>';
            }
        });
        $('#divisionFilter').html(divHtml);
    } else {
        var divHtml = '<option value="">All Divisions</option>';
        state.divisions.forEach(function(d) {
            divHtml += '<option value="' + d.id + '">' + d.name + '</option>';
        });
        $('#divisionFilter').html(divHtml);
    }

    searchInvolvements();
}

function renderInvolvementList(involvements) {
    if (!involvements || involvements.length === 0) {
        $('#involvementList').html('<div class="ip-empty-state py-4"><i class="bi bi-inbox d-block"></i><span>No involvements found</span></div>');
        return;
    }

    var html = '';
    involvements.forEach(function(inv) {
        var selected = state.sourceOrgId == inv.id ? ' ip-selected' : '';
        html += '<div class="ip-involvement-item' + selected + '" data-id="' + inv.id + '" data-name="' + inv.name.replace(/"/g, '&quot;') + '" onclick="selectInvolvement(' + inv.id + ', this)">';
        html += '<div class="ip-org-name">' + inv.name + '</div>';
        html += '<div class="ip-org-meta">' + inv.program + ' / ' + inv.division + ' &bull; ' + inv.memberCount + ' members</div>';
        html += '</div>';
    });
    $('#involvementList').html(html);
}

function selectInvolvement(orgId, el) {
    // Update UI
    $('#involvementList .ip-involvement-item').removeClass('ip-selected');
    $(el).addClass('ip-selected');

    var orgName = $(el).data('name');
    state.sourceOrgId = orgId;
    state.sourceOrgName = orgName;

    $('#selectedOrgName').text(orgName + ' (#' + orgId + ')');
    $('#selectedInvolvement').show();

    // Reset effort state for new source involvement
    state.currentEffort = null;
    state.processed = [];
    state.effortDirty = false;

    // Load registrants, questions, and efforts
    loadRegistrants(orgId);
    loadEfforts(orgId);
    loadWatchlist(orgId);

    // Enable watchlist button
    $('#watchlistBtn').prop('disabled', false);

    // If "same involvement" is selected for target, update target info
    if ($('#targetSame').is(':checked')) {
        state.targetOrgId = orgId;
        updateTargetCounts();
    }
}

function clearSourceSelection() {
    state.sourceOrgId = null;
    state.sourceOrgName = '';
    state.registrants = [];
    state.questions = [];
    state.subgroups = [];
    state.queue = [];

    // Clear effort state
    clearEffortState();

    $('#selectedInvolvement').hide();
    $('#involvementList .ip-involvement-item').removeClass('ip-selected');
    $('#registrantList').html('<div class="ip-empty-state"><i class="bi bi-inbox d-block"></i><span>Select a source involvement to load registrants</span></div>');
    $('#questionToggles').html('<div class="text-muted small">Select an involvement to see questions</div>');
    $('#subgroupList').html('<div class="text-muted small">Select source involvement first</div>');
    renderQueue();

    $('#watchlistBtn').prop('disabled', true);
}

// ============================================================================
// Load Registrants
// ============================================================================
function loadRegistrants(orgId) {
    $('#registrantLoading').show();

    ajax('load_registrants', { org_id: orgId }, function(response) {
        $('#registrantLoading').hide();
        console.log('load_registrants response:', response);

        if (response.success) {
            // Ensure arrays are always initialized (defensive coding)
            state.registrants = response.members || [];
            state.questions = response.questions || [];
            state.subgroups = response.subgroups || [];

            console.log('Loaded:', state.registrants.length, 'registrants,', state.questions.length, 'questions,', state.subgroups.length, 'subgroups');

            renderQuestionToggles();
            populateFilterDropdowns();
            clearAllFilters();  // Reset filters when loading new involvement
            renderRegistrantList();

            // Don't call renderSubgroups() here - the target subgroups are managed by:
            // - loadTargetSubgroups() when "Same involvement" mode is active (called from selectInvolvement)
            // - loadTargetSubgroups() when user selects a target in "Different involvement" mode

            var memberCount = state.registrants.length;
            $('#registrantCount').text(memberCount + ' people');
        } else {
            showToast(response.message || 'Error loading registrants', 'danger');
        }
    });
}

function renderQuestionToggles() {
    if (!state.questions || state.questions.length === 0) {
        $('#questionToggles').html('<div class="text-muted small">No registration questions found</div>');
        return;
    }

    var html = '';
    state.questions.forEach(function(q, idx) {
        var checked = state.selectedQuestions.indexOf(q.id) >= 0 ? ' checked' : '';
        html += '<div class="ip-question-toggle">';
        html += '<input type="checkbox" id="q_' + q.id + '" data-qid="' + q.id + '"' + checked + '>';
        html += '<label for="q_' + q.id + '">' + q.text + '</label>';
        html += '</div>';
    });
    $('#questionToggles').html(html);
}

function toggleQuestion(qid) {
    var idx = state.selectedQuestions.indexOf(qid);
    if (idx >= 0) {
        state.selectedQuestions.splice(idx, 1);
    } else {
        state.selectedQuestions.push(qid);
    }
    renderRegistrantList();
}

function renderRegistrantList() {
    // Ensure state.registrants is an array
    if (!state.registrants || !Array.isArray(state.registrants)) {
        state.registrants = [];
    }

    var filters = getActiveFilters();
    var filterMode = 'all';
    if ($('#filterPending').hasClass('btn-light')) filterMode = 'pending';
    if ($('#filterProcessed').hasClass('btn-light')) filterMode = 'processed';

    var filtered = state.registrants.filter(function(r) {
        // Name filter
        if (filters.name && r.name.toLowerCase().indexOf(filters.name) < 0) return false;

        // Status filter (pending/processed)
        var procFlag = isProcessed(r.peopleId);
        if (filterMode === 'pending' && procFlag) return false;
        if (filterMode === 'processed' && !procFlag) return false;

        // Gender filter
        if (filters.gender && r.gender !== filters.gender) return false;

        // Age filter
        if (filters.ageMin !== null || filters.ageMax !== null) {
            var ageValue;
            if (filters.ageUnit === 'months') {
                // Use age in months (for young children like Mother's Day Out)
                ageValue = r.ageMonths;
            } else {
                // Use age in years
                ageValue = r.age;
            }

            // Skip if no age data available
            if (ageValue === null || ageValue === undefined) return false;

            if (filters.ageMin !== null && ageValue < filters.ageMin) return false;
            if (filters.ageMax !== null && ageValue > filters.ageMax) return false;
        }

        // Subgroup filter
        if (filters.subgroup) {
            if (!r.subgroups || r.subgroups.indexOf(filters.subgroup) < 0) return false;
        }

        // Question/Answer filter
        if (filters.question && filters.answer) {
            var answerValue = r.answers && r.answers[filters.question];
            if (!answerValue || answerValue.toLowerCase().indexOf(filters.answer) < 0) return false;
        }

        return true;
    });

    // Update count display
    var totalCount = state.registrants.length;
    var filteredCount = filtered.length;
    if (filteredCount < totalCount) {
        $('#registrantCount').text(filteredCount + ' of ' + totalCount + ' shown');
    } else {
        $('#registrantCount').text(totalCount + ' people');
    }

    if (filtered.length === 0) {
        $('#registrantList').html('<div class="ip-empty-state"><i class="bi bi-inbox d-block"></i><span>No matching registrants</span></div>');
        return;
    }

    var html = '';
    filtered.forEach(function(r) {
        var inQueue = state.queue.some(function(q) { return q.peopleId === r.peopleId; });
        var procFlag = isProcessed(r.peopleId);
        var procInfo = getProcessedInfo(r.peopleId);

        var cardClass = 'ip-person-card';
        if (inQueue) cardClass += ' ip-selected';
        if (procFlag) cardClass += ' ip-processed';

        html += '<div class="' + cardClass + '" data-pid="' + r.peopleId + '" onclick="toggleQueuePerson(' + r.peopleId + ')">';
        html += '<div class="ip-name">' + r.name;
        if (procFlag) html += ' <i class="bi bi-check-circle-fill text-success"></i>';
        html += '</div>';

        // Show assignment info for processed people
        if (procFlag && procInfo && procInfo.targetOrgName) {
            html += '<div class="ip-assignment-info">';
            html += '<i class="bi bi-arrow-right-circle me-1"></i>';
            html += '<span class="text-success">' + escapeHtml(procInfo.targetOrgName) + '</span>';
            if (procInfo.subgroups && procInfo.subgroups.length > 0) {
                html += ' <i class="bi bi-tag-fill ms-1 me-1"></i>';
                html += '<span class="text-info">' + procInfo.subgroups.map(escapeHtml).join(', ') + '</span>';
            }
            html += '</div>';
        }
        html += '<div class="ip-details">';
        // Show age - in months if filter is set to months, otherwise in years
        if (filters.ageUnit === 'months' && r.ageMonths !== null && r.ageMonths !== undefined) {
            html += r.ageMonths + ' mo &bull; ';
        } else if (r.age) {
            html += r.age + ' yrs &bull; ';
        }
        if (r.gender) html += r.gender + ' &bull; ';
        html += r.memberType;
        html += '</div>';

        // Show source involvement subgroups (if any)
        if (r.subgroups && r.subgroups.length > 0) {
            html += '<div class="ip-subgroups small text-muted">';
            html += '<i class="bi bi-tag me-1"></i>' + r.subgroups.map(escapeHtml).join(', ');
            html += '</div>';
        }

        // Show selected question answers
        // r.answers is an object {questionText: answer, ...}
        // state.selectedQuestions contains question IDs
        // state.questions contains {id, text} mappings
        if (state.selectedQuestions.length > 0 && r.answers && typeof r.answers === 'object') {
            var answerKeys = Object.keys(r.answers);
            if (answerKeys.length > 0) {
                // Get the selected question texts
                var selectedQuestionTexts = state.questions
                    .filter(function(q) { return state.selectedQuestions.indexOf(q.id) >= 0; })
                    .map(function(q) { return q.text; });

                html += '<div class="ip-answers">';
                var hasAnswers = false;
                answerKeys.forEach(function(questionText) {
                    // Only show if this question is selected (or show all if we can't match)
                    if (selectedQuestionTexts.length === 0 || selectedQuestionTexts.indexOf(questionText) >= 0) {
                        var answer = r.answers[questionText];
                        if (answer) {
                            html += '<div class="ip-answer-item"><strong>' + questionText + ':</strong> ' + answer + '</div>';
                            hasAnswers = true;
                        }
                    }
                });
                if (!hasAnswers) {
                    html += '<div class="ip-answer-item text-muted">No answers for selected questions</div>';
                }
                html += '</div>';
            }
        }

        html += '</div>';
    });

    $('#registrantList').html(html);
}

function filterRegistrantList() {
    renderRegistrantList();
}

function filterRegistrants(mode) {
    $('#filterAll, #filterPending, #filterProcessed').removeClass('btn-light').addClass('btn-outline-light');
    if (mode === 'all') $('#filterAll').removeClass('btn-outline-light').addClass('btn-light');
    if (mode === 'pending') $('#filterPending').removeClass('btn-outline-light').addClass('btn-light');
    if (mode === 'processed') $('#filterProcessed').removeClass('btn-outline-light').addClass('btn-light');

    renderRegistrantList();
}

// ============================================================================
// Advanced Filtering
// ============================================================================
function toggleAdvancedFilters() {
    var panel = $('#advancedFiltersPanel');
    var toggle = $('#advancedFiltersToggle');
    if (panel.is(':visible')) {
        panel.slideUp(200);
        toggle.html('<i class="bi bi-funnel me-1"></i>Show Advanced Filters');
    } else {
        panel.slideDown(200);
        toggle.html('<i class="bi bi-funnel-fill me-1"></i>Hide Advanced Filters');
    }
}

function populateFilterDropdowns() {
    // Populate subgroup filter
    var sgHtml = '<option value="">All Subgroups</option>';
    if (state.subgroups && state.subgroups.length > 0) {
        state.subgroups.forEach(function(sg) {
            sgHtml += '<option value="' + escapeHtml(sg.name) + '">' + escapeHtml(sg.name) + ' (' + sg.count + ')</option>';
        });
    }
    $('#filterSubgroup').html(sgHtml);

    // Populate question filter (truncate long questions for display)
    var qHtml = '<option value="">Select question...</option>';
    if (state.questions && state.questions.length > 0) {
        state.questions.forEach(function(q) {
            var displayText = q.text;
            if (displayText && displayText.length > 60) {
                displayText = displayText.substring(0, 57) + '...';
            }
            // Value keeps full text, display is truncated
            qHtml += '<option value="' + escapeHtml(q.text) + '">' + escapeHtml(displayText) + '</option>';
        });
    }
    $('#filterQuestion').html(qHtml);
}

function updateAnswerFilter() {
    var selectedQuestion = $('#filterQuestion').val();
    if (selectedQuestion) {
        $('#answerFilterContainer').show();
    } else {
        $('#answerFilterContainer').hide();
        $('#filterAnswer').val('');
    }
    applyFilters();
}

function applyFilters() {
    renderRegistrantList();
}

function clearAllFilters() {
    $('#registrantSearch').val('');
    $('#filterGender').val('');
    $('#filterAgeUnit').val('years');
    $('#filterAgeMin').val('');
    $('#filterAgeMax').val('');
    $('#filterSubgroup').val('');
    $('#filterQuestion').val('');
    $('#filterAnswer').val('');
    $('#answerFilterContainer').hide();
    renderRegistrantList();
}

function getActiveFilters() {
    var searchVal = $('#registrantSearch').val();
    var answerVal = $('#filterAnswer').val();

    return {
        name: searchVal ? searchVal.toLowerCase() : '',
        gender: $('#filterGender').val() || '',
        ageUnit: $('#filterAgeUnit').val() || 'years',
        ageMin: $('#filterAgeMin').val() ? parseInt($('#filterAgeMin').val()) : null,
        ageMax: $('#filterAgeMax').val() ? parseInt($('#filterAgeMax').val()) : null,
        subgroup: $('#filterSubgroup').val() || '',
        question: $('#filterQuestion').val() || '',
        answer: answerVal ? answerVal.toLowerCase() : ''
    };
}

// ============================================================================
// Selection Queue
// ============================================================================
function toggleQueuePerson(peopleId) {
    var person = state.registrants.find(function(r) { return r.peopleId === peopleId; });
    if (!person) return;

    var idx = state.queue.findIndex(function(q) { return q.peopleId === peopleId; });
    if (idx >= 0) {
        state.queue.splice(idx, 1);
    } else {
        state.queue.push(person);
    }

    renderQueue();
    renderRegistrantList();
    updateProcessButton();
    updateEffortDisplay();
}

function addToQueueFromSearch(peopleId, name, source) {
    // Check if already in queue
    if (state.queue.some(function(q) { return q.peopleId === peopleId; })) {
        showToast('Already in queue', 'warning');
        return;
    }

    state.queue.push({
        peopleId: peopleId,
        name: name,
        source: source
    });

    renderQueue();
    renderRegistrantList();
    updateProcessButton();
    updateEffortDisplay();
    showToast(name + ' added to queue', 'success');
}

function removeFromQueue(peopleId) {
    state.queue = state.queue.filter(function(q) { return q.peopleId !== peopleId; });
    renderQueue();
    renderRegistrantList();
    updateProcessButton();
    updateEffortDisplay();
}

function clearQueue() {
    state.queue = [];
    renderQueue();
    renderRegistrantList();
    updateProcessButton();
    updateEffortDisplay();
}

function renderQueue() {
    if (state.queue.length === 0) {
        $('#queueList').html('<div class="ip-empty-state py-3"><i class="bi bi-plus-circle d-block" style="font-size: 24px;"></i><small>Click on registrants to add them</small></div>');
        $('#queueActions').hide();
        $('#queueCount').text('0');
        return;
    }

    var html = '';
    state.queue.forEach(function(q) {
        html += '<div class="ip-queue-item">';
        html += '<span class="ip-name">' + q.name + '</span>';
        html += '<span class="ip-remove-btn" onclick="event.stopPropagation(); removeFromQueue(' + q.peopleId + ')"><i class="bi bi-x-lg"></i></span>';
        html += '</div>';
    });

    $('#queueList').html(html);
    $('#queueActions').show();
    $('#queueCount').text(state.queue.length);
}

// ============================================================================
// Match Finding
// ============================================================================
function searchMembers() {
    var term = $('#memberSearch').val();
    if (!state.sourceOrgId) {
        $('#memberResults').html('<div class="text-muted small text-center py-3">Select source involvement first</div>');
        return;
    }

    if (!term || term.length < 2) {
        $('#memberResults').html('<div class="text-muted small text-center py-3">Type at least 2 characters</div>');
        return;
    }

    ajax('search_members', { org_id: state.sourceOrgId, search_term: term }, function(response) {
        if (response.success) {
            renderMemberResults(response.members);
        }
    });
}

function renderMemberResults(members) {
    if (!members || members.length === 0) {
        $('#memberResults').html('<div class="text-muted small text-center py-3">No members found</div>');
        return;
    }

    var html = '';
    members.forEach(function(m) {
        var inQueue = state.queue.some(function(q) { return q.peopleId === m.peopleId; });
        html += '<div class="ip-person-card' + (inQueue ? ' ip-selected' : '') + '" data-people-id="' + m.peopleId + '" data-name="' + escapeHtml(m.name) + '" data-source="member">';
        html += '<div class="ip-name">' + m.name + '</div>';
        html += '<div class="ip-details">' + (m.age || '') + ' ' + (m.gender || '') + '</div>';
        html += '</div>';
    });
    $('#memberResults').html(html);

    // Bind click handlers using jQuery (avoids quote escaping issues)
    $('#memberResults .ip-person-card').off('click').on('click', function() {
        var peopleId = $(this).data('people-id');
        var name = $(this).data('name');
        var source = $(this).data('source');
        addToQueueFromSearch(peopleId, name, source);
    });
}

function searchGlobalPeople() {
    var term = $('#globalSearch').val();

    if (!term || term.length < 2) {
        $('#globalResults').html('<div class="text-muted small text-center py-3">Type at least 2 characters</div>');
        return;
    }

    ajax('search_people', { search_term: term }, function(response) {
        if (response.success) {
            renderGlobalResults(response.people);
        }
    });
}

function renderGlobalResults(people) {
    if (!people || people.length === 0) {
        $('#globalResults').html('<div class="text-muted small text-center py-3">No people found</div>');
        return;
    }

    var html = '';
    people.forEach(function(p) {
        var inQueue = state.queue.some(function(q) { return q.peopleId === p.peopleId; });
        html += '<div class="ip-person-card' + (inQueue ? ' ip-selected' : '') + '" data-people-id="' + p.peopleId + '" data-name="' + escapeHtml(p.name) + '" data-source="global">';
        html += '<div class="ip-name">' + p.name + '</div>';
        html += '<div class="ip-details">' + (p.email || '') + ' &bull; ' + (p.status || '') + '</div>';
        html += '</div>';
    });
    $('#globalResults').html(html);

    // Bind click handlers using jQuery (avoids quote escaping issues)
    $('#globalResults .ip-person-card').off('click').on('click', function() {
        var peopleId = $(this).data('people-id');
        var name = $(this).data('name');
        var source = $(this).data('source');
        addToQueueFromSearch(peopleId, name, source);
    });
}

// ============================================================================
// Target Configuration
// ============================================================================
function updateTargetMode() {
    if ($('#targetDifferent').is(':checked')) {
        $('#targetOrgSection').show();
        state.targetOrgId = null;
        state.targetSubgroups = [];
        $('#subgroupList').html('<div class="text-muted small">Select a target involvement first</div>');
        $('#countsDisplay').html('<div class="text-muted small">Select target to see counts</div>');
    } else {
        $('#targetOrgSection').hide();
        state.targetOrgId = state.sourceOrgId;
        updateTargetCounts();
    }
    updateProcessButton();
}

function searchTargetOrg() {
    var term = $('#targetOrgSearch').val();

    if (!term || term.length < 2) {
        $('#targetOrgList').html('<div class="text-muted small text-center py-3">Type to search</div>');
        return;
    }

    ajax('search_involvements', { search_term: term }, function(response) {
        if (response.success) {
            var html = '';
            response.involvements.forEach(function(inv) {
                var selected = state.targetOrgId == inv.id ? ' ip-selected' : '';
                html += '<div class="ip-involvement-item' + selected + '" data-org-id="' + inv.id + '" data-org-name="' + escapeHtml(inv.name) + '">';
                html += '<div class="ip-org-name">' + inv.name + '</div>';
                html += '<div class="ip-org-meta">' + inv.memberCount + ' members</div>';
                html += '</div>';
            });
            $('#targetOrgList').html(html);

            // Bind click handlers using jQuery (avoids quote escaping issues)
            $('#targetOrgList .ip-involvement-item').off('click').on('click', function() {
                var orgId = $(this).data('org-id');
                var orgName = $(this).data('org-name');
                selectTargetOrg(orgId, orgName);
            });
        }
    });
}

function selectTargetOrg(orgId, orgName) {
    state.targetOrgId = orgId;
    state.targetOrgName = orgName;
    $('#targetOrgList .ip-involvement-item').removeClass('ip-selected');
    $('#targetOrgList .ip-involvement-item[data-org-id="' + orgId + '"]').addClass('ip-selected');

    // Load subgroups and counts for the selected target
    loadTargetSubgroups(orgId);
    updateProcessButton();
}

function loadTargetSubgroups(orgId) {
    $('#subgroupList').html('<div class="text-center py-2"><div class="spinner-border spinner-border-sm"></div> Loading subgroups...</div>');
    $('#countsDisplay').html('<div class="text-center py-2"><div class="spinner-border spinner-border-sm"></div></div>');

    ajax('get_target_counts', { org_ids: orgId }, function(response) {
        console.log('loadTargetSubgroups response:', response);
        console.log('loadTargetSubgroups targets[0]:', response.targets ? response.targets[0] : 'no targets');
        console.log('loadTargetSubgroups subgroups:', response.targets && response.targets[0] ? response.targets[0].subgroups : 'no subgroups');
        if (response.success && response.targets && response.targets.length > 0) {
            var target = response.targets[0];

            // Update counts display
            var countsHtml = '<div class="ip-count-badge"><strong>Total Members:</strong>&nbsp;' + target.totalCount + '</div>';
            $('#countsDisplay').html(countsHtml);

            // Update subgroups list with checkboxes and counts
            if (target.subgroups && target.subgroups.length > 0) {
                state.targetSubgroups = target.subgroups;
                var html = '';
                target.subgroups.forEach(function(sg) {
                    console.log('Rendering subgroup:', sg.name, 'count:', sg.count);
                    var checked = state.selectedSubgroups.indexOf(sg.name) >= 0 ? ' checked' : '';
                    html += '<div class="form-check">';
                    html += '<input class="form-check-input" type="checkbox" id="tsg_' + sg.name.replace(/\\s/g, '_') + '" data-sgname="' + sg.name.replace(/"/g, '&quot;') + '"' + checked + '>';
                    html += '<label class="form-check-label" for="tsg_' + sg.name.replace(/\\s/g, '_') + '">';
                    html += sg.name + ' <span class="badge bg-secondary">' + sg.count + '</span>';
                    html += '</label>';
                    html += '</div>';
                });
                console.log('Setting #subgroupList HTML, length:', html.length);
                $('#subgroupList').html(html);
                console.log('Subgroups rendered, #subgroupList content length:', $('#subgroupList').html().length);
            } else {
                state.targetSubgroups = [];
                $('#subgroupList').html('<div class="text-muted small">No subgroups defined for this involvement</div>');
            }
        } else {
            console.log('loadTargetSubgroups failed, falling back to state.subgroups');
            // Fallback: If target counts failed, use subgroups from loadRegistrants if available
            // This handles the case where get_target_counts has issues but load_registrants worked
            if (state.subgroups && state.subgroups.length > 0 && state.targetOrgId === state.sourceOrgId) {
                state.targetSubgroups = state.subgroups;
                renderSubgroups();
                $('#countsDisplay').html('<div class="text-muted small">Counts not available</div>');
            } else {
                $('#countsDisplay').html('<div class="text-muted small">Could not load counts</div>');
                $('#subgroupList').html('<div class="text-muted small">No subgroups found</div>');
            }
        }
    });
}

function renderSubgroups() {
    if (!state.subgroups || state.subgroups.length === 0) {
        $('#subgroupList').html('<div class="text-muted small">No subgroups defined</div>');
        return;
    }

    var html = '';
    state.subgroups.forEach(function(sg) {
        var checked = state.selectedSubgroups.indexOf(sg.name) >= 0 ? ' checked' : '';
        html += '<div class="form-check">';
        html += '<input class="form-check-input" type="checkbox" id="sg_' + sg.name.replace(/\s/g, '_') + '" data-sgname="' + sg.name.replace(/"/g, '&quot;') + '"' + checked + '>';
        html += '<label class="form-check-label" for="sg_' + sg.name.replace(/\s/g, '_') + '">' + sg.name + ' <span class="badge bg-secondary">' + sg.count + '</span></label>';
        html += '</div>';
    });
    $('#subgroupList').html(html);
}

function toggleSubgroup(sgName) {
    var idx = state.selectedSubgroups.indexOf(sgName);
    if (idx >= 0) {
        state.selectedSubgroups.splice(idx, 1);
    } else {
        state.selectedSubgroups.push(sgName);
    }
    updateProcessButton();
}

function createSubgroup() {
    var name = $('#newSubgroup').val().trim();
    if (!name) return;

    // Add to local state
    state.subgroups.push({ name: name, count: 0 });
    state.selectedSubgroups.push(name);

    renderSubgroups();
    $('#newSubgroup').val('');
    showToast('Subgroup "' + name + '" will be created when processing', 'info');
}

function updateTargetCounts() {
    // Refresh subgroups and counts for the current target
    console.log('updateTargetCounts called, targetOrgId:', state.targetOrgId);
    if (state.targetOrgId) {
        loadTargetSubgroups(state.targetOrgId);
    } else {
        $('#countsDisplay').html('<div class="text-muted small">Select target to see counts</div>');
        $('#subgroupList').html('<div class="text-muted small">Select a target involvement first</div>');
    }
}

// ============================================================================
// Process Queue
// ============================================================================
function updateProcessButton() {
    var canProcess = state.queue.length > 0 && state.targetOrgId;
    $('#processBtn').prop('disabled', !canProcess);
}

function processQueue() {
    if (!state.targetOrgId || state.queue.length === 0) {
        showToast('Select target and add people to queue', 'warning');
        return;
    }

    var peopleIds = state.queue.map(function(q) { return q.peopleId; }).join(',');
    var subgroups = state.selectedSubgroups.join(',');

    $('#processBtn').prop('disabled', true).html('<span class="spinner-border spinner-border-sm me-2"></span>Processing...');

    ajax('process_queue', {
        people_ids: peopleIds,
        target_org_id: state.targetOrgId,
        subgroup_names: subgroups
    }, function(response) {
        $('#processBtn').prop('disabled', false).html('<i class="bi bi-check2-circle me-2"></i>Process Queue');

        if (response.success) {
            showToast(response.message, 'success');

            // Determine target name
            var targetName = state.targetOrgName || 'Unknown';
            if ($('#targetSame').is(':checked') && state.sourceOrgName) {
                targetName = state.sourceOrgName + ' (same)';
            }

            // Mark as processed in local state with assignment details
            state.queue.forEach(function(q) {
                addProcessedPerson(q.peopleId, state.targetOrgId, targetName, state.selectedSubgroups.slice());
            });

            // If we have a current effort, save the processed list to the effort
            if (state.currentEffort) {
                state.effortDirty = true;
                // Auto-save the effort with updated processed list
                autoSaveEffort();
            } else {
                // No effort - save to global processed list (legacy behavior)
                state.queue.forEach(function(q) {
                    ajax('mark_processed', { org_id: state.sourceOrgId, people_id: q.peopleId });
                });
            }

            // Clear queue
            state.queue = [];
            renderQueue();
            renderRegistrantList();
            updateTargetCounts();
            updateProcessButton();
            updateEffortDisplay();

            // Refresh target org search results to show updated member counts
            if ($('#targetOrgSearch').val()) {
                searchTargetOrg();
            }

            // Refresh subgroup counts
            if (state.sourceOrgId) {
                loadRegistrants(state.sourceOrgId);
            }
        } else {
            showToast('Error: ' + response.message, 'danger');
        }
    });
}

function autoSaveEffort() {
    console.log('autoSaveEffort called, currentEffort:', state.currentEffort, 'sourceOrgId:', state.sourceOrgId);
    if (!state.currentEffort || !state.sourceOrgId) {
        console.log('autoSaveEffort early return - no currentEffort or sourceOrgId');
        return;
    }

    var config = {
        selectedQuestions: state.selectedQuestions,
        selectedSubgroups: state.selectedSubgroups
    };

    console.log('Saving effort with processed:', state.processed);
    ajax('save_effort', {
        org_id: state.sourceOrgId,
        effort_id: state.currentEffort.id,
        effort_name: state.currentEffort.name,
        effort_config: JSON.stringify(config),
        processed_json: getProcessedJson()
    }, function(response) {
        console.log('save_effort response:', response);
        if (response.success) {
            state.effortDirty = false;
            updateEffortDisplay();
        } else {
            console.error('save_effort failed:', response.message);
        }
    });
}

// ============================================================================
// Processed List Management
// ============================================================================
function loadProcessedList(orgId) {
    ajax('get_processed', { org_id: orgId }, function(response) {
        if (response.success) {
            state.processed = response.processed || [];
            renderRegistrantList();
        }
    });
}

// ============================================================================
// Configuration Save/Load
// ============================================================================
function saveCurrentConfig() {
    if (!state.sourceOrgId) {
        showToast('Select an involvement first', 'warning');
        return;
    }

    var config = {
        selectedQuestions: state.selectedQuestions,
        selectedSubgroups: state.selectedSubgroups,
        targetSameOrg: $('#targetSame').is(':checked'),
        savedAt: new Date().toISOString()
    };

    ajax('save_config', {
        org_id: state.sourceOrgId,
        config_data: JSON.stringify(config)
    }, function(response) {
        if (response.success) {
            showToast('Configuration saved', 'success');
        } else {
            showToast('Error saving configuration', 'danger');
        }
    });
}

function loadLastConfig() {
    if (!state.sourceOrgId) {
        showToast('Select an involvement first', 'warning');
        return;
    }

    ajax('load_config', { org_id: state.sourceOrgId }, function(response) {
        if (response.success && response.config) {
            var config = response.config;

            // Apply config
            state.selectedQuestions = config.selectedQuestions || [];
            state.selectedSubgroups = config.selectedSubgroups || [];

            if (config.targetSameOrg) {
                $('#targetSame').prop('checked', true);
            } else {
                $('#targetDifferent').prop('checked', true);
            }

            renderQuestionToggles();
            renderSubgroups();
            updateTargetMode();
            renderRegistrantList();

            showToast('Configuration loaded (saved ' + new Date(config.savedAt).toLocaleDateString() + ')', 'success');
        } else {
            showToast('No saved configuration found', 'info');
        }
    });
}

// ============================================================================
// Watchlist
// ============================================================================
function loadWatchlist(orgId) {
    ajax('get_watchlist', { org_id: orgId }, function(response) {
        if (response.success) {
            state.watchlist = response.watchlist || [];
            renderWatchlist();
        }
    });
}

function renderWatchlist() {
    if (!state.watchlist || state.watchlist.length === 0) {
        $('#watchlistItems').html('<div class="ip-empty-state py-3"><small>No pending matches</small></div>');
        $('#watchlistCount').text('0');
        return;
    }

    var html = '';
    state.watchlist.forEach(function(item, idx) {
        html += '<div class="ip-queue-item">';
        html += '<div>';
        html += '<div class="ip-name">Waiting: #' + item.watcherId + '</div>';
        html += '<small class="text-muted">Looking for: ' + item.requestedName + '</small>';
        html += '</div>';
        html += '<span class="ip-remove-btn" onclick="removeWatchlistItem(' + idx + ')"><i class="bi bi-x-lg"></i></span>';
        html += '</div>';
    });

    $('#watchlistItems').html(html);
    $('#watchlistCount').text(state.watchlist.length);
}

function addToWatchlist() {
    if (!state.sourceOrgId) {
        showToast('Select an involvement first', 'warning');
        return;
    }

    // Populate the watcher dropdown with queue items
    var html = '<option value="">Select from queue...</option>';
    state.queue.forEach(function(q) {
        html += '<option value="' + q.peopleId + '">' + q.name + '</option>';
    });

    // Also add from registrants
    state.registrants.forEach(function(r) {
        if (!state.queue.some(function(q) { return q.peopleId === r.peopleId; })) {
            html += '<option value="' + r.peopleId + '">' + r.name + '</option>';
        }
    });

    $('#watchlistWatcher').html(html);
    $('#watchlistRequestedName').val('');

    var modal = new bootstrap.Modal(document.getElementById('watchlistModal'));
    modal.show();
}

function saveWatchlistItem() {
    var watcherId = $('#watchlistWatcher').val();
    var requestedName = $('#watchlistRequestedName').val().trim();

    if (!watcherId || !requestedName) {
        showToast('Please fill in all fields', 'warning');
        return;
    }

    ajax('add_watchlist', {
        org_id: state.sourceOrgId,
        watcher_id: watcherId,
        requested_name: requestedName
    }, function(response) {
        if (response.success) {
            state.watchlist = response.watchlist;
            renderWatchlist();
            bootstrap.Modal.getInstance(document.getElementById('watchlistModal')).hide();
            showToast('Added to watchlist', 'success');
        } else {
            showToast('Error: ' + response.message, 'danger');
        }
    });
}

function removeWatchlistItem(idx) {
    ajax('remove_watchlist', { org_id: state.sourceOrgId, index: idx }, function(response) {
        if (response.success) {
            state.watchlist = response.watchlist;
            renderWatchlist();
            showToast('Removed from watchlist', 'success');
        }
    });
}

// ============================================================================
// Help
// ============================================================================
function showHelp() {
    var modal = new bootstrap.Modal(document.getElementById('helpModal'));
    modal.show();
}
</script>

<!-- Debug output div - shows JS errors visually -->
<div id="jsDebugOutput" class="alert alert-danger m-3" style="display: none;"></div>

<div class="ip-root ip-main-container">
    <!-- Header -->
    <div class="d-flex justify-content-between align-items-center mb-4">
        <div>
            <h2 class="mb-1"><i class="bi bi-people-fill me-2"></i>Involvement Processor</h2>
            <p class="text-muted mb-0">Process registrations, match pairs, and manage track assignments</p>
        </div>
        <div>
            <button class="btn btn-outline-secondary me-2" onclick="showHelp()">
                <i class="bi bi-question-circle"></i> Help
            </button>
        </div>
    </div>

    <div class="row">
        <!-- Left Column: Source & Registrants -->
        <div class="col-lg-8">
            <!-- Setup Panel -->
            <div class="ip-panel ip-setup-panel">
                <div class="ip-panel-header">
                    <span><i class="bi bi-gear-fill me-2"></i>Setup</span>
                    <div>
                        <button class="btn btn-sm btn-light me-1" onclick="loadLastConfig()">
                            <i class="bi bi-clock-history"></i> Last Used
                        </button>
                        <button class="btn btn-sm btn-light" onclick="saveCurrentConfig()">
                            <i class="bi bi-save"></i> Save Setup
                        </button>
                    </div>
                </div>
                <div class="ip-panel-body">
                    <div class="row">
                        <div class="col-md-6">
                            <label class="form-label fw-bold">Source Involvement</label>
                            <div class="ip-search-box mb-2">
                                <i class="bi bi-search ip-search-icon"></i>
                                <input type="text" id="involvementSearch" class="form-control"
                                       placeholder="Search by name or ID...">
                            </div>
                            <div class="row mb-2">
                                <div class="col-6">
                                    <select id="programFilter" class="form-select form-select-sm">
                                        <option value="">All Programs</option>
                                    </select>
                                </div>
                                <div class="col-6">
                                    <select id="divisionFilter" class="form-select form-select-sm">
                                        <option value="">All Divisions</option>
                                    </select>
                                </div>
                            </div>
                            <div id="involvementList" style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; border-radius: 6px;">
                                <div class="ip-empty-state py-4">
                                    <i class="bi bi-search d-block"></i>
                                    <span>Search for an involvement</span>
                                </div>
                            </div>
                            <div id="selectedInvolvement" class="mt-2" style="display: none;">
                                <div class="alert alert-info mb-0 py-2">
                                    <strong>Selected:</strong> <span id="selectedOrgName"></span>
                                    <button type="button" class="btn-close float-end" onclick="clearSourceSelection()"></button>
                                </div>
                            </div>

                            <!-- Effort Management -->
                            <div id="effortPanel" class="mt-3" style="display: none;">
                                <div class="card border-primary">
                                    <div class="card-body py-2 px-3">
                                        <div class="d-flex justify-content-between align-items-center">
                                            <div>
                                                <i class="bi bi-bookmark-star text-primary me-1"></i>
                                                <span id="effortName" class="fw-bold small">No effort selected</span>
                                                <span id="effortDirtyIndicator" class="badge bg-warning text-dark ms-1" style="display: none;">unsaved</span>
                                            </div>
                                            <div class="btn-group btn-group-sm">
                                                <button class="btn btn-outline-primary btn-sm" type="button" onclick="showEffortPicker()" title="Switch effort">
                                                    <i class="bi bi-list"></i>
                                                </button>
                                                <button class="btn btn-outline-primary btn-sm" type="button" onclick="renameCurrentEffort()" title="Rename effort" id="effortRenameBtn">
                                                    <i class="bi bi-pencil"></i>
                                                </button>
                                            </div>
                                        </div>
                                        <div id="effortProgress" class="small text-muted mt-1"></div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label fw-bold">Display Questions</label>
                            <div id="questionToggles" style="max-height: 250px; overflow-y: auto;">
                                <div class="text-muted small">Select an involvement to see questions</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Registrants Panel -->
            <div class="ip-panel">
                <div class="ip-panel-header">
                    <span><i class="bi bi-person-lines-fill me-2"></i>Registrants</span>
                    <div>
                        <span class="badge bg-light text-dark" id="registrantCount">0 people</span>
                        <div class="btn-group btn-group-sm ms-2">
                            <button class="btn btn-light" onclick="filterRegistrants('all')" id="filterAll">All</button>
                            <button class="btn btn-outline-light" onclick="filterRegistrants('pending')" id="filterPending">Pending</button>
                            <button class="btn btn-outline-light" onclick="filterRegistrants('processed')" id="filterProcessed">Processed</button>
                        </div>
                    </div>
                </div>
                <div class="ip-panel-body" style="position: relative;">
                    <!-- Basic search -->
                    <div class="ip-search-box mb-2">
                        <i class="bi bi-search ip-search-icon"></i>
                        <input type="text" id="registrantSearch" class="form-control"
                               placeholder="Search by name...">
                    </div>

                    <!-- Advanced Filters Toggle -->
                    <div class="mb-2">
                        <a href="#" class="small text-primary" onclick="toggleAdvancedFilters(); return false;" id="advancedFiltersToggle">
                            <i class="bi bi-funnel me-1"></i>Show Advanced Filters
                        </a>
                    </div>

                    <!-- Advanced Filters Panel -->
                    <div id="advancedFiltersPanel" style="display: none;" class="mb-3 p-2 bg-light rounded">
                        <div class="row g-2">
                            <!-- Gender Filter -->
                            <div class="col-6 col-md-3">
                                <label class="form-label small mb-1">Gender</label>
                                <select id="filterGender" class="form-select form-select-sm" onchange="applyFilters()">
                                    <option value="">All</option>
                                    <option value="M">Male</option>
                                    <option value="F">Female</option>
                                </select>
                            </div>

                            <!-- Age Filter Mode -->
                            <div class="col-6 col-md-3">
                                <label class="form-label small mb-1">Age Unit</label>
                                <select id="filterAgeUnit" class="form-select form-select-sm" onchange="applyFilters()">
                                    <option value="years">Years</option>
                                    <option value="months">Months</option>
                                </select>
                            </div>

                            <!-- Age Range -->
                            <div class="col-6 col-md-3">
                                <label class="form-label small mb-1">Age Min</label>
                                <input type="number" id="filterAgeMin" class="form-control form-control-sm"
                                       placeholder="Min" min="0" onchange="applyFilters()">
                            </div>
                            <div class="col-6 col-md-3">
                                <label class="form-label small mb-1">Age Max</label>
                                <input type="number" id="filterAgeMax" class="form-control form-control-sm"
                                       placeholder="Max" min="0" onchange="applyFilters()">
                            </div>

                            <!-- Subgroup Filter -->
                            <div class="col-12 col-md-6">
                                <label class="form-label small mb-1">Subgroup</label>
                                <select id="filterSubgroup" class="form-select form-select-sm" onchange="applyFilters()">
                                    <option value="">All Subgroups</option>
                                    <!-- Populated dynamically -->
                                </select>
                            </div>

                            <!-- Question/Answer Filter -->
                            <div class="col-12 col-md-6">
                                <label class="form-label small mb-1">Question</label>
                                <select id="filterQuestion" class="form-select form-select-sm" onchange="updateAnswerFilter()">
                                    <option value="">Select question...</option>
                                    <!-- Populated dynamically -->
                                </select>
                            </div>
                            <div class="col-12" id="answerFilterContainer" style="display: none;">
                                <label class="form-label small mb-1">Answer Contains</label>
                                <input type="text" id="filterAnswer" class="form-control form-control-sm"
                                       placeholder="Type to filter by answer..." onkeyup="applyFilters()">
                            </div>
                        </div>
                        <div class="mt-2 text-end">
                            <button class="btn btn-sm btn-outline-secondary" onclick="clearAllFilters()">
                                <i class="bi bi-x-circle me-1"></i>Clear Filters
                            </button>
                        </div>
                    </div>

                    <div id="registrantList" class="ip-registrant-grid">
                        <div class="ip-empty-state">
                            <i class="bi bi-inbox d-block"></i>
                            <span>Select a source involvement to load registrants</span>
                        </div>
                    </div>
                    <div id="registrantLoading" class="ip-loading-overlay" style="display: none;">
                        <div class="spinner-border text-primary" role="status">
                            <span class="visually-hidden">Loading...</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Right Column: Queue, Target, Watchlist -->
        <div class="col-lg-4">
            <!-- Selection Queue -->
            <div class="ip-panel ip-queue-panel ip-queue-sticky">
                <div class="ip-panel-header">
                    <span><i class="bi bi-collection me-2"></i>Selection Queue</span>
                    <span class="badge" id="queueCount">0</span>
                </div>
                <div class="ip-panel-body">
                    <div id="queueList">
                        <div class="ip-empty-state py-3">
                            <i class="bi bi-plus-circle d-block" style="font-size: 24px;"></i>
                            <small>Click on registrants to add them</small>
                        </div>
                    </div>
                    <div id="queueActions" style="display: none;">
                        <hr>
                        <button class="btn btn-outline-danger btn-sm" onclick="clearQueue()">
                            <i class="bi bi-trash"></i> Clear All
                        </button>
                    </div>
                </div>
            </div>

            <!-- Match Finder - Hidden for now, not fully functional -->
            <!-- TODO: Enable this when match finding feature is ready
            <div class="ip-panel">
                <div class="ip-panel-header">
                    <span><i class="bi bi-link-45deg me-2"></i>Find Match</span>
                </div>
                <div class="ip-panel-body">
                    <ul class="nav nav-tabs mb-3" role="tablist">
                        <li class="nav-item">
                            <a class="nav-link active" data-bs-toggle="tab" href="#sameOrgTab">Same Involvement</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" data-bs-toggle="tab" href="#globalTab">Global Search</a>
                        </li>
                    </ul>
                    <div class="tab-content">
                        <div class="tab-pane fade show active" id="sameOrgTab">
                            <div class="ip-search-box mb-2">
                                <i class="bi bi-search ip-search-icon"></i>
                                <input type="text" id="memberSearch" class="form-control form-control-sm"
                                       placeholder="Search members...">
                            </div>
                            <div id="memberResults" style="max-height: 200px; overflow-y: auto;">
                                <div class="text-muted small text-center py-3">Type to search members</div>
                            </div>
                        </div>
                        <div class="tab-pane fade" id="globalTab">
                            <div class="ip-search-box mb-2">
                                <i class="bi bi-search ip-search-icon"></i>
                                <input type="text" id="globalSearch" class="form-control form-control-sm"
                                       placeholder="Search all people...">
                            </div>
                            <div id="globalResults" style="max-height: 200px; overflow-y: auto;">
                                <div class="text-muted small text-center py-3">Type to search people</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            -->

            <!-- Target Configuration -->
            <div class="ip-panel ip-target-panel">
                <div class="ip-panel-header">
                    <span><i class="bi bi-bullseye me-2"></i>Target Assignment</span>
                </div>
                <div class="ip-panel-body">
                    <div class="form-check mb-2">
                        <input class="form-check-input" type="radio" name="targetType" id="targetSame" checked>
                        <label class="form-check-label" for="targetSame">Same involvement + Subgroup</label>
                    </div>
                    <div class="form-check mb-3">
                        <input class="form-check-input" type="radio" name="targetType" id="targetDifferent">
                        <label class="form-check-label" for="targetDifferent">Different involvement</label>
                    </div>

                    <div id="targetOrgSection" style="display: none;">
                        <label class="form-label small fw-bold">Target Involvement</label>
                        <div class="ip-search-box mb-2">
                            <i class="bi bi-search ip-search-icon"></i>
                            <input type="text" id="targetOrgSearch" class="form-control form-control-sm"
                                   placeholder="Search...">
                        </div>
                        <div id="targetOrgList" style="max-height: 150px; overflow-y: auto; border: 1px solid #ddd; border-radius: 6px;">
                        </div>
                    </div>

                    <div id="subgroupSection" class="mt-3">
                        <label class="form-label small fw-bold">Subgroups</label>
                        <div id="subgroupList">
                            <div class="text-muted small">Select source involvement first</div>
                        </div>
                        <div class="mt-2">
                            <input type="text" id="newSubgroup" class="form-control form-control-sm"
                                   placeholder="Create new subgroup...">
                            <button class="btn btn-outline-secondary btn-sm mt-1" onclick="createSubgroup()">
                                <i class="bi bi-plus"></i> Add Subgroup
                            </button>
                        </div>
                    </div>

                    <div id="targetCounts" class="mt-3">
                        <label class="form-label small fw-bold">Current Counts</label>
                        <div id="countsDisplay">
                            <div class="text-muted small">Select target to see counts</div>
                        </div>
                    </div>

                    <hr>

                    <button class="btn btn-success ip-action-btn w-100" onclick="processQueue()" id="processBtn" disabled>
                        <i class="bi bi-check2-circle me-2"></i>Process Queue
                    </button>
                </div>
            </div>

            <!-- Watchlist - Hidden for now, not fully utilized
            <div class="ip-panel ip-watchlist-panel">
                <div class="ip-panel-header">
                    <span><i class="bi bi-eye me-2"></i>Watchlist</span>
                    <span class="badge" id="watchlistCount">0</span>
                </div>
                <div class="ip-panel-body">
                    <div id="watchlistItems">
                        <div class="ip-empty-state py-3">
                            <small>No pending matches</small>
                        </div>
                    </div>
                    <div class="mt-2">
                        <button class="btn btn-outline-danger btn-sm w-100" onclick="addToWatchlist()" id="watchlistBtn" disabled>
                            <i class="bi bi-plus-circle"></i> Add Pending Match
                        </button>
                    </div>
                </div>
            </div>
            -->
        </div>
    </div>

    <!-- Add to Watchlist Modal - Hidden for now
<div class="modal fade" id="watchlistModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Add to Watchlist</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="mb-3">
                    <label class="form-label">Waiting Person</label>
                    <select id="watchlistWatcher" class="form-select">
                        <option value="">Select from queue...</option>
                    </select>
                </div>
                <div class="mb-3">
                    <label class="form-label">Requested Match Name</label>
                    <input type="text" id="watchlistRequestedName" class="form-control"
                           placeholder="Name of person they want to match with...">
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button type="button" class="btn btn-primary" onclick="saveWatchlistItem()">Add to Watchlist</button>
            </div>
        </div>
    </div>
</div>
-->

<!-- Help Modal -->
<div class="modal fade" id="helpModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Involvement Processor Help</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <h6>Use Case 1: Pair Matching (Bunk-mates, Roommates)</h6>
                <ol>
                    <li>Select your source involvement (e.g., Summer Camp)</li>
                    <li>Enable the question(s) showing pair preferences</li>
                    <li>Click on a registrant to add them to the queue</li>
                    <li>Look at the assignment info to see where their requested match was placed</li>
                    <li>Add other people to the queue who should go together</li>
                    <li>Select target subgroup(s) and click Process</li>
                </ol>

                <h6 class="mt-4">Use Case 2: Track/Session Assignment with Load Balancing</h6>
                <ol>
                    <li>Select your source involvement</li>
                    <li>Enable the question showing track preferences</li>
                    <li>View current counts in each subgroup</li>
                    <li>Select people and assign to appropriate tracks</li>
                    <li>Counts update automatically after processing</li>
                </ol>

            </div>
        </div>
    </div>
</div>

<!-- Continue Previous Work Modal -->
<div class="modal fade" id="continueEffortModal" tabindex="-1" data-bs-backdrop="static">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header bg-primary text-white">
                <h5 class="modal-title"><i class="bi bi-bookmark-check me-2"></i>Previous Work Found</h5>
            </div>
            <div class="modal-body">
                <p>You have saved work for this involvement:</p>
                <div id="previousEffortsList" class="list-group mb-3">
                    <!-- Populated dynamically -->
                </div>
                <p class="text-muted small mb-0">Choose an effort to continue, or start fresh with a new effort.</p>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-outline-secondary" onclick="startFreshEffort()">
                    <i class="bi bi-plus-circle me-1"></i>Start Fresh
                </button>
            </div>
        </div>
    </div>
</div>

<!-- New Effort Modal -->
<div class="modal fade" id="newEffortModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title"><i class="bi bi-bookmark-plus me-2"></i>Name Your Effort</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="mb-3">
                    <label class="form-label">Effort Name</label>
                    <input type="text" id="newEffortName" class="form-control"
                           placeholder="e.g., Cabin Assignments, Bus Groups, Track Selection...">
                    <div class="form-text">Give this effort a descriptive name to identify it later.</div>
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                <button type="button" class="btn btn-primary" onclick="createNewEffort()">
                    <i class="bi bi-plus-circle me-1"></i>Create Effort
                </button>
            </div>
        </div>
    </div>
</div>

</div><!-- End ip-root -->

<!-- Toast Container -->
<div class="ip-toast-container" id="toastContainer"></div>
</body>
</html>
'''
    # Output for both PyScript (print) and PyScriptForm (model.Form)
    print html
    model.Form = html
