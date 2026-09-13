# /api/v1 endpoints I call

**Scope:** `/api/v1` **only.** Anything not on that path is in Part 3 so you know it exists, not as a request.

---

## Part 1. Complete list, and where each one is covered

Every /api/v1 endpoint any of my five apps touches, grouped by need. "Extra permission" means anything beyond being a valid login; blank means open to any authenticated caller.

### 1.1 Works today with a token, nothing needed (22)


| Verb      | Route                                          | Extra permission | Used by                            |
| --------- | ---------------------------------------------- | ---------------- | ---------------------------------- |
| POST      | `/v1/Account/CreateUserAccessToken`            | Access           | all five, the token exchange       |
| GET       | `/v1/SpecialContentFiles`                      | see note         | DisplayCache, scheduler, Scan, Go  |
| POST      | `/v1/SpecialContent`                           | see note         | DisplayCache, scheduler, Scan, Go  |
| PUT       | `/v1/SpecialContent/{id}`                      | see note         | DisplayCache, scheduler, Scan, Go  |
| GET, POST | `/v1/Integration/Contacts`                     |                  | TPxi Go                            |
| POST      | `/v1/Integration/Contacts/{peopleId}/Calls`    |                  | TPxi Go                            |
| POST      | `/v1/Integration/Contacts/{peopleId}/Messages` |                  | TPxi Go                            |
| PUT       | `/v1/Integration/Contacts/{peopleId}/Phone`    |                  | TPxi Go                            |
| GET       | `/v1/People/{peopleId}/Involvements`           |                  | TPxi Go                            |
| GET       | `/v1/People/{peopleId}/Extended`               |                  | TPxi Go, see also 2c               |
| GET       | `/v1/People/{peopleId}/Profile`                |                  | TPxi Go                            |
| GET       | `/v1/People/{peopleId}/TaskNotes/Assigned`     |                  | TPxi Go                            |
| GET       | `/v1/People/GetUserProfile`                    |                  | TPxi Go                            |
| GET       | `/v1/lookup/MemberStatuses`                    |                  | TPxi Go                            |
| GET       | `/v1/Keywords`                                 |                  | TPxi Go, incl. nested extra values |
| POST      | `/v1/TaskNotes/Create`                         | Access           | TPxi Go                            |
| POST      | `/v1/TaskNotes/{id}/Accept, /Decline`          | Access           | TPxi Go                            |
| POST      | `/v2/TaskNotes/{id}/Complete`                  | Access           | TPxi Go                            |
| PUT       | `/v1/TaskNotes/{id}`                           | Access           | TaskNote tool                      |
| GET, POST | `/v1/Tags/* (all four)`                        |                  | TPxi Go                            |
| POST      | `/v1/Messages/EmailQueue`                      |                  | queued email, see 2e               |
| GET       | `/v1/Messages/EmailQueue/{id}/Status`          |                  | queued email                       |


**The Special Content note.** These are PAT enabled and the attribute carries
no role. The restriction is in the service underneath: `CanRetrieveSpecialContent`
wants **Admin or SpecialContentFull** for any content type, **SpecialContentBasic**
for Text and Html only, and **EmailTemplates** for the template types. My
installs write a Python script, which is its own type, so they need Admin or
SpecialContentFull. That is fine for me, and I think requiring a high privilege
to write executable Python is right. It bites ordinary users instead, which is
2d rather than a complaint here.

### 1.2 Rooms and Reservations, the one that still needs Basic (1)


| Verb | Route                            | Extra permission | Used by                 |
| ---- | -------------------------------- | ---------------- | ----------------------- |
| GET  | `/v1/Reservables/Locations/Tree` |                  | DisplayCache, scheduler |


This is the only endpoint I call today that is neither PAT enabled nor
anonymous, so it is the one thing on your API still holding Basic open for me.

**On today's usage I can work around it, and I will.** In the conference
scheduler it is not fetching rooms, it is a connectivity canary in a `probe()`
function that even treats 403 as a pass, and in DisplayCache it is a fallback
behind a Python script. Repointing the probe at `/v1/SpecialContentFiles` is a
one line change. So this does not block your Basic cutover and I am not asking
you to treat it as urgent.

**But I do not think the right fix is me routing around it**, and this is the
one place in this document where I am asking about direction rather than a
flag. When I went to check, `Reservables` is **0 of 34 endpoints PAT enabled**.
`Reservations` next to it is **0 of 42**. With `Meetings`, that is 77 endpoints
across the rooms and booking area with no token access at all.

That matters to me because rooms and reservations is where I expect to build
next. Room booking, resource scheduling and the teaching and class scheduling
work I want to do all live in that area, and right now none of it is reachable
from anything but a browser session. A one line probe change fixes my
yesterday. It does not change the fact that an entire module is closed.

So the question is not really "please enable this endpoint", it is **is
Reservables and Reservations something you intend to open to tokens, or is it
deliberately browser only?** If it is coming, I will wait and plan around it.
If it is not, I would rather know now, before I design something that assumes
it.

Some context for why I am asking rather than assuming: across the whole API
**101 of 677 endpoints are PAT enabled, about 15 percent**, and 41 of the 58
modules have none at all. I am not suggesting all of those should change, most
of them are clearly internal. But it does mean I cannot tell the difference
between "not enabled yet" and "never going to be" by looking, which is why the
list in Part 2 is phrased as requests rather than assumptions.

### 1.3 Covered by a request below (18)


| Verb     | Route                                                       | Section | Note                                 |
| -------- | ----------------------------------------------------------- | ------- | ------------------------------------ |
| GET      | `/v1/People/Search/{criteria}`                              | 1b      | paging, not PAT, see the correction  |
| GET      | `/v1/People/Search/Autocomplete`                            | 1b      | type ahead                           |
| GET      | `/v1/Attendance/Organizations`                              | 1b      | also 2b                              |
| GET      | `/v1/Organizations/{orgId}`                                 | 1b      | also 2b                              |
| GET      | `/v1/Directory/People/{peopleId}/Organizations`             | 1b      | also 2a                              |
| GET      | `/v1/Attendance/Organizations/{orgId}/Meetings`             | 2b      | care workflow                        |
| GET      | `/v1/Attendance/Organizations/{orgId}/Meetings/{meetingId}` | 2b      | care workflow                        |
| GET      | `/v1/Attendance/Meetings/Available`                         | 2b      | care workflow                        |
| GET      | `/v1/Attendance/Meetings/{meetingId}/Members`               | 2b      | see the Members/Guests note          |
| GET      | `/v1/Attendance/Meetings/{meetingId}/Guests`                | 2b      | see the Members/Guests note          |
| GET      | `/v1/Involvements/{orgId}/Meetings`                         | 2b      | upcoming meetings, role gated        |
| GET      | `/v1/People/{peopleId}/WeeklyEngagementScores`              | 2a      | person scoped, exists today          |
| (new)    | `a person scoped attendance read`                           | 2a      | the main gap                         |
| (new)    | `members of an involvement`                                 | 2a      | the main gap                         |
| (varies) | `extra value writes`                                        | 2c      |                                      |
| (varies) | `Special Content for ordinary users`                        | 2d      |                                      |
| POST     | `/v1/Messages/EmailQueue/Process`                           | 2e      | swagger says PAT, attribute does not |
| POST     | `/v1/Messages/EmailQueue/{id}/Send`                         | 2e      | not in swagger at all                |




### 1.4 Their own auth scheme, no action (1)


| Verb | Route             | Extra permission       | Used by              |
| ---- | ----------------- | ---------------------- | -------------------- |
| GET  | `/v1/SitesData/*` | `Authorization: Token` | conference scheduler |




### 1.5 Two I have raised separately (2)

`GET /v1/Contents` and `GET /v1/lookup/Campuses` are both `[AnonymousEndpoint]`
with no `[FunctionAuthorize]`, so the role middleware runs them with no
authentication at all. I had both on my earlier list as needing PAT and that
was wrong, they need nothing. I am not treating them as API requests. I have
sent you a separate note about `/v1/Contents` specifically.

---



## Part 1b. Still outstanding from the list I sent before

**20 of the 25 I sent last time are already PAT enabled**, 
so thank you, and I have stopped asking about those. These five are
the remainder, and they are all the same one line change.


| Verb | Route                                           | Note                                                         |
| ---- | ----------------------------------------------- | ------------------------------------------------------------ |
| GET  | `/v1/People/Search/{criteria}`                  | Full search, but **see the note** below about the result cap |
| GET  | `/v1/People/Search/Autocomplete`                | Type ahead                                                   |
| GET  | `/v1/Attendance/Organizations`                  | Also in 2b below                                             |
| GET  | `/v1/Organizations/{orgId}`                     | Also in 2b below                                             |
| GET  | `/v1/Directory/People/{peopleId}/Organizations` | Also in 2a below                                             |


**One correction to what I asked for last time.** I said I needed
`People/Search/{criteria}` because `Integration/Contacts` caps at 10 results.
Looking at it properly, that endpoint has the same cap:

```
v1/People/Search/{criteria}        Search(criteria, 1, 10)
v1/Integration/Contacts            SearchIntegrationContacts(query, userId, 1, 10)
```

Both hardcode page 1 and 10 results. So PAT enabling `People/Search` on its own
would not actually help me, it would just be a second endpoint that returns ten
rows.

The underlying services already page. `SearchIntegrationContacts` is declared
as `(string searchExpression, int currentUserId, int? page = 1, int? take = 10)`
and the endpoint simply never passes them through. The same `(1, 10)` is
hardcoded at four call sites, three in `People_SearchOperations.cs` and one in
`Integration_ContactOperations.cs`, which makes it look more like an oversight
that propagated than a deliberate limit.

So what I actually need is **page and take exposed as query parameters on the
search endpoint**, on whichever of the two you would rather I used. That is a
smaller change than PAT enabling a second search, and it is the thing that
unblocks me.

**It is worth saying how I work around this today, because the workaround is
the argument.** On the phone apps I do not use your search at all. A Python
script walks the directory, chunks it 2000 contacts at a time into Special
Content with a manifest, and the app pulls every chunk on sync and loads them
into a local SQLite database that it then searches itself. That is why the
mobile apps can find the fortieth Smith and the browser ones cannot.

It works, and I am not asking you to replace it, since offline lookup is worth
having on its own for caller ID. But it does mean I am using Special Content as
a bulk export channel for the whole church directory, which is not what it is
for, and I am doing that because the search endpoint will not page. Ten rows is
also a quiet failure rather than a loud one: nothing in the response says the
list was truncated, so staff searching a common surname see ten results and
reasonably conclude the eleventh person is not in the database.

---



## Part 2. Wish list



### If you only do one thing

**A way to list the members of one involvement with a token.** Everything else
in Part 2 is an improvement to something that already works. This one is the
difference between a feature shipping and not shipping, and it is the smallest
item on the list.

The thing waiting on it is our pastoral and hospital care report, which is next
up for the mobile app and has been in the queue a while. I went through it
endpoint by endpoint, and **the roster is the only piece a token cannot reach**:


| What the report does                                        | How                                                                 | Status      |
| ----------------------------------------------------------- | ------------------------------------------------------------------- | ----------- |
| Read the care list roster                                   | members of one involvement                                          | **blocked** |
| Read a person's visit notes, with keywords and extra values | `/v1/People/{peopleId}/TaskNotes`                                   | works       |
| Log a visit with facility, room and so on                   | `/v1/TaskNotes/Create`, `TaskNoteDto` carries `TaskNoteExtraValues` | works       |
| Populate the facility picker                                | `/v1/Keywords`, nested extra values and options                     | works       |
| Close out a visit                                           | `/v2/TaskNotes/{id}/Complete`                                       | works       |
| Add and remove people from the care list                    | `Join` and `Leave`                                                  | works       |
| Show who is on call this week                               | `Organizations/{orgId}/RollList`, which carries `commitmentId`      | works       |
| Person detail and member status                             | `/Profile`, `/lookup/MemberStatuses`                                | works       |


Eight of nine already work with a PAT, which is genuinely good, and most of
that is thanks to the batch you enabled last round. The ninth is a roster read,
and without it the app cannot answer "who are we visiting" and none of the rest
is reachable from a phone.

I am not asking for the full person level attendance read in 2a to get this
done. That one matters for the prospect work and I would still like it, but it
is a bigger conversation. **The roster on its own unblocks pastoral care**, and
if it is easier to scope, I would take it as its own change ahead of everything
else here.

One note on `RollList`, since I am relying on it above. It is PAT enabled and
returns what I need, but its own description says it creates a meeting when one
does not exist for that date, so it is a write. That is fine for our on call
schedule, where the meetings are already built in advance. I mention it only
because a read shaped endpoint that writes is easy to call by accident, and
`Meetings/{meetingId}/Members` in 2b would be the clean read if it were open.

### 2a. Members of an involvement, and a person's attendance

These two are the real gap, and they are related.

I can already get a person's involvements from
`GET /v1/People/{peopleId}/Involvements`, which works fine with a token. What I
cannot do from there is either of the obvious next steps:

**Members of one of those involvements.** There is no safe read for this:

- `/v1/Attendance/Meetings/{meetingId}/Members` is per meeting rather than per
involvement, and is not PAT enabled
- `/v1/Attendance/Organizations/{orgId}/RollList` is PAT enabled, but it is a
POST that creates a meeting if none exists for that date. Your own
description says it is therefore a write operation, so I cannot use it to
read a roster without writing to the church's data
- `/v1/DataWarehouse/InvolvementMembers` is bulk and gated on Developer +
ApiOnly

**A person's attendance history.** Every attendance route is scoped to a
meeting or an organization. Nothing returns one person's attendance. The only
person level attendance data anywhere is `/v1/DataWarehouse/Attendance`, same
gate.

So: **a PAT can already record attendance for a person and list what
involvements they belong to, but cannot read the members of those involvements
or whether the person attended anything.**

For the attendance one I am deliberately not proposing a route or parameters,
since you will know how it should fit the rest of the API. The requirements are
only:

- scoped to one person or preferably a list of people so it can be a single call
- honors whatever the calling user is already entitled to see
- can return the full history or a date bounded slice
- paged, some people have thousands of rows

`POST /v1/Resources/Search` **is already this shape** Its own description 
says it returns a paged, filtered list where "access is scoped per caller 
using campus, organization membership/type, and status flags", and it is PAT 
enabled. That entitlement model is exactly what I am asking for: scoped to 
what the calling user can already see, rather than gated behind a blanket 
high privilege role. If a person attendance read followed the same pattern 
it would solve this completely.

One more that is already the right shape and only needs the flag:


| Verb | Route                                          | Note                        |
| ---- | ---------------------------------------------- | --------------------------- |
| GET  | `/v1/People/{peopleId}/WeeklyEngagementScores` | Person scoped, exists today |




### 2b. Attendance reads for the care workflow (5)

The Attendance module is 27 endpoints: 17 already allow PAT, and every write is
enabled. The gap is reads. A token can create a meeting through `GetOrCreate`
and record attendance to it, but cannot read the meeting back.


| Verb | Route                                                       |
| ---- | ----------------------------------------------------------- |
| GET  | `/v1/Attendance/Organizations/{orgId}/Meetings`             |
| GET  | `/v1/Attendance/Organizations/{orgId}/Meetings/{meetingId}` |
| GET  | `/v1/Attendance/Meetings/Available`                         |
| GET  | `/v1/Attendance/Meetings/{meetingId}/Members`               |
| GET  | `/v1/Attendance/Meetings/{meetingId}/Guests`                |


(`/v1/Attendance/Organizations` and `/v1/Organizations/{orgId}` belong here too
but are already on the 1b list, so I have not repeated them.)

**One thing to flag about Members and Guests.** Those two endpoints do not
split the way the names suggest, and between them they may not give me
everybody who attended. Reading the queries, `Members` excludes
`MemberTypeCodes.Prospect`, and `Guests` returns people who are not non
prospect members. So a prospect who attends comes back as a guest, and the real
split is "prospect or unrecognized" against "everyone else" rather than member
against visitor.

**Prospects are the reason this matters most to me**, and it is worth
explaining why rather than just asking. Our prospect tool scores engagement as
a ladder, and every rung on it is an attendance type:


| Rung | Attendance type          | Meaning               |
| ---- | ------------------------ | --------------------- |
| 5    | 10, leader               | leading something     |
| 4    | serving types, 70 and 20 | serving somewhere     |
| 3    | 30, member               | attending as a member |
| 2    | 50, recent guest         | been back             |
| 1    | 60, new guest            | first time            |


That ladder is the whole product. Someone moving from 1 to 2 is the single
most important event we track, and someone falling from 3 to nothing is the
second. A two bucket split flattens rungs 5, 4 and 3 into "members" and 2 and 1
into "guests", which erases every transition worth acting on, and prospects
fall out of the member bucket entirely.

I do not think any of this needs new endpoints, because the data is already
there:

- `AttendancePersonDto` already carries `MemberTypeId` and `AttendMemberTypeId`,
so if the PAT enabled responses include those fields I can do the rest myself
- there is already an `attendeeType` parameter on
`GetAttendanceMeetingMembers` (0 or omitted for both, 1 for members, 2 for
guests) that the grid path uses but these two endpoints do not expose

So the ask is: PAT enable both, and if it is cheap, let a caller get every
attendee in one call with the attendance and member type ids on each row,
rather than having to guess which of two buckets someone landed in.

The same point applies to the person level attendance read in 2a. If it
returns the attendance type id per row, that one endpoint plus these two would
let the prospect work move to mobile as is.

This completes moving a person between care statuses from the app: hospital,
recovery, homebound, nursing home. `Join` and `Leave` already allow PAT.

**Upcoming meetings would be a great add.** `Meetings/Available` and
`Organizations/{orgId}/Meetings` may already cover it depending on how they
filter. There is also `GET /v1/Involvements/{orgId}/Meetings`, which looks like
exactly the right thing but is gated on Admin, Checkin or Edit and is not PAT
enabled. Any one of the three would answer "what is scheduled next for this
involvement", which is the common question on a phone, rather than what
already happened.

### 2c. Extra values

Extra values are where churches put everything TouchPoint does not model
natively, and a lot of my tooling stores state there. Through the API today:


| Need                           | Status                                        |
| ------------------------------ | --------------------------------------------- |
| Read person extra values       | `/v1/People/{peopleId}/Extended`, PAT enabled |
| Write person extra values      | **no endpoint anywhere**                      |
| Read involvement extra values  | **no endpoint anywhere**                      |
| Write involvement extra values | **no endpoint anywhere**                      |


The only Extras endpoints are `Organizations/Meetings/Extras`, which are
meeting extras rather than involvement ones, and none allow PAT.

### 2d. Special Content writes for ordinary users

A migration problem rather than a permissions complaint, worth flagging before
the direction is settled.

`model.WriteContentText` in Python has **no role check at all**, so any 
user who can run a script can store state in Special Content. That is how 
per user state works across a lot of my tools.  TPxi Dashboards, for instance, 
stores each user's dashboards that way.

The REST equivalent runs `CanCreateSpecialContent`, which needs Admin,
SpecialContentFull, or SpecialContentBasic for Text and Html. So moving that
code from Python to REST is not a like for like port for ordinary staff. On my
database that is 69 users with Basic, 26 with Admin or Full, and everyone else
loses the ability to save their own settings.

I am not asking you to weaken the API check. Writing executable Python should
need a high privilege and I would not want SpecialContentFull handed out. But
if the direction is REST first, there needs to be some supported way for a
normal user to persist their own state, or a lot of per user features stop
working when they move.

### 2e. Two smaller questions

**Sending queued email, and a swagger mismatch.** I went and read how the
queue actually drains, and I think this is closer to a documentation question
than a request.

`EnqueueEmail` sets `ReadyToSend = scheduleTime != null`, and the background
picker requires `ReadyToSend == true`. So:

- if I create a queue **with** a `scheduleTime`, it is marked ready, the
background job picks it up, and the mail goes out. A PAT can already do this
end to end, since create and status are both enabled
- if I create one **without** a `scheduleTime`, it is left not ready, and
something has to call `Process` or `{id}/Send` to move it. Neither of those is
PAT enabled

So a token can send email today as long as it always sets a schedule time, and
I am happy to just do that.

On `Process` specifically, I checked whether the swagger entry was just a
documentation slip and it is not. The attribute is the actual gate:

- `FunctionAuthorizeAttribute` defaults `AllowPATAccess` to **false**
- `TouchPointAuthenticationMiddleware` only enters the PAT branch when that
flag is true. On a plain `[FunctionAuthorize]` the branch is skipped, so the
token is never looked up and `UserId` stays 0
- `ProcessEmailQueue` then calls `CanManageEmailQueue(0)`, which fails, and
returns **403 "You do not have permission to process email queues."**

So swagger lists `POST /v1/Messages/EmailQueue/Process` under the note saying
these endpoints take a PAT, but a PAT gets a 403 there. Worth fixing in one
direction or the other. The message is the part that cost me time: it points at
roles, when the real cause is that PAT authentication was never attempted, so
anyone hitting this goes off checking their own permissions first.

`POST /v1/Messages/EmailQueue/{id}/Send` is the same plain attribute and is not
in swagger at all.

If always passing a schedule time is the intended pattern for API callers, say
so and I will treat that as the answer and stop asking for the other two.

**Creating a tag by name.** I cannot find one. The four Tags endpoints all take
a `tagId` that must already exist, and `InitializeSessionTempTag` creates a
temporary one. Checking I have not missed something.  This would be highly useful 
for fast follow-up and persistence in TPxiGo.

---



### 2f. Managing involvement membership

Adding, moving and removing people from involvements is the other thing I
cannot do from an app, and I use it constantly on the Python side. Across my
scripts: `JoinOrg` 30 times, `DropOrgMember` 26, `SetMemberType` 24,
`AddSubGroup` 33, `MoveToOrg` 7.

The API has two of these. `Join` and `Leave` are both PAT enabled and need no
extra role, which is great. The gap is that `Join` is not really "join", it is
"join as a Member":

```csharp
// AttendanceService.JoinOrganization
var orgMemberDto = await _dataStoreService.AddOrganizationMember(
    peopleId, organizationId, (int)MemberTypeCodes.Member);
```

The member type is hardcoded. It is not that the plumbing cannot carry one,
the layer directly underneath takes it and the service simply does not pass it
through:

```
IDataStoreService_Organizations
    AddOrganizationMember(int peopleId, int organizationId, int memberTypeId)

IAttendanceService
    JoinOrganization(int organizationId, int peopleId, int currentUserId, bool isPatCaller)
                                                        no memberTypeId
```

This is the same shape as the search cap in 1b, where the service takes `page`
and `take` and the endpoint hardcodes `(1, 10)`. In both cases the capability
is already there one level down and only the public surface is fixed.

So from an app I can add
someone, but never as a Leader, a Prospect, a Volunteer or anything else, and I
cannot change the type of someone already there. There is no endpoint for that
anywhere, and none for moving a person between involvements.

What I would ask for, smallest first:


| Need                             | Suggestion                                           |
| -------------------------------- | ---------------------------------------------------- |
| Add with a chosen member type    | an optional `memberTypeId` on the existing `Join`    |
| Change an existing member's type | the equivalent of `SetMemberType`                    |
| Move between involvements        | the equivalent of `MoveToOrg`                        |
| Sub groups                       | the equivalent of `AddSubGroup` and `RemoveSubGroup` |


The first one is the one I would take if only one were possible, because it is
an optional parameter on an endpoint that already exists and already has the
right permission model.

**This is also what makes the prospect ladder in 2b actionable rather than just
observable.** That whole tool is built on member type. Converting someone means
promoting them off the prospect type onto a real one, and that is the moment
the work pays off. Today a token can read the ladder but cannot move anyone up
it, so the app can tell a staff member what to do and then cannot do it. Doing
the promotion is the point.

Worth saying plainly: `Leave` today is a drop with no way to put the person
back as anything other than a Member, so "move" implemented as Leave then Join
loses both the member type and the enrollment date. I would rather not build
that workaround.

## Part 3. Not /api/v1, listed so you know rather than as requests

- `/PythonApi/{name}` for #API scripts. Lives in CmsWeb, which has no PAT
support at all; `AuthHelper.AuthenticateDeveloper` base64 decodes the
Authorization header unconditionally, so only Basic can reach it. It already
requires Developer + APIOnly. This is why I still hold a password even after
everything else moves to PAT: DisplayCache pulls its data through
`/pythonapi/` every 6 hours per church.
- `/api/Organizations` with OData query params, used by the conference
scheduler. Under `/api/` but not `/api/v1/`, so I am assuming out of scope.
Say if not.
- `/v1/SitesData/*` feeds, used by the scheduler. These use
`Authorization: Token <sitesToken>`, so neither Basic nor PAT. No change
needed unless you are consolidating schemes.
- `/Portrait/{peopleId}` for photos in DisplayCache. Not v1.

---



## Summary

**Nothing here blocks you removing Basic from /api/v1.** The single endpoint I
still call that needs it, `Reservables/Locations/Tree`, is a connectivity probe
I can repoint in one line, and I will.

**Two things I need answers on rather than code changes:**

1. **How does a church get its first token once Basic is gone?** Minting one

through `CreateUserAccessToken` currently starts from a username and password.
2. **Are Reservables and Reservations ever going to accept tokens?** That is 76
endpoints with none today, and it is where I expect to build next. I am not
asking you to do it now, I am asking whether to plan around it.

Ranked by what actually stops me building something today:

1. **A roster read for one involvement.** Smallest item here and the only one

that decides whether something ships. Pastoral and hospital care is otherwise
finished against endpoints that already work, eight of nine, and is waiting on
this alone. If it is easier to scope on its own, please do that ahead of the
rest of this document.
2. **2a, a person scoped attendance read.** The other half of 2a and a bigger
conversation. This is what the prospect work needs, and it wants the attendance
type id on each row.
3. **2b, five attendance reads**, plus returning the attendance and member type
ids so the Members and Guests split does not flatten prospects into guests.
4. **1b, page and take on search.** Smaller than what I asked for last time, and
it replaces that ask rather than adding to it.
5. **2c and 2d, extra values and Special Content for ordinary users.** These are
about what non admin staff can do from Python, not about my apps.
6. **2e, two housekeeping items**, one of which is a swagger and attribute
disagreement rather than a request.

Things I listed before that turned out to need nothing, so you can ignore them:
`Contents` and `lookup/Campuses` are anonymous, the Tags endpoints and the
Special Content writes are already enabled, and the email queue already sends
through a PAT as long as the caller sets a schedule time.
