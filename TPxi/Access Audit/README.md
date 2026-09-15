
# Access Audit

When somebody leaves your staff, can you say in one place everything they could
still get to? This answers that, and then lets you clean it up.

It is also the reference nobody has the rest of the year: who holds which role,
and what that role actually lets them do.

## Installing

One file, one upload.

1. Go to Admin, Advanced, Special Content, Python
2. Add New, name it `TPxi_AccessAudit`, paste in `TPxi_AccessAudit.py`, save
3. In Special Content, Text Content, CustomReports, add this line:

```xml
<Report name="TPxi_AccessAudit" type="PyScript" role="Admin" />
```

That is it. Nothing to schedule, nothing to keep up to date. It reads your
database live every time you open it, and it tells you when a newer version is
available.

Admins only. Everything on the page is staff names, logins and roles.

## What the four tabs show

**By role.** Pick a role and see who holds it, what it unlocks, and what it
changes on screen. Useful before you hand somebody a role, and useful when you
are wondering why two people see different things.

**By person.** Pick a person and see everything attached to them: their logins,
their roles, the groups they lead, the registration emails that come to them,
their open tasks, their access tokens, and what they have been doing lately.
Work down the page and you have offboarded them.

**Tokens.** Access tokens let an outside app into your database without a
password. This lists every live one, who owns it, when it expires, and whether
anything has actually used it in the past year.

**What to do.** The findings pulled together as a worklist, with the things you
can be sure about at the top.

## What you can change from here

Nothing changes until you tick some boxes and confirm.

| | Sends emails? | Can you undo it? |
| --- | --- | --- |
| Reassign somebody's open tasks | **yes, one per task** | reassign them back |
| Mark them complete | **yes, one per task** | no |
| Archive them | no | yes |
| Delete them | no | **no** |
| Remove a role | no | add it back |
| Revoke an access token | no | **no** |

**Watch the emails.** Reassigning or completing a task sends a notification for
every single one. One person here had 3,701 open tasks, which would have been
3,701 emails. TouchPoint gives no way to turn that off, so the tool tells you
the number before you confirm. If you just want the list cleared quietly,
archive instead. Archived tasks are still there and can be brought back.

**Revoking tokens is switched off until you turn it on.** Ask your TouchPoint
admin to add a setting called `TPxi_AccessAudit_PAT`. Without it you can still
see every token, you just cannot delete one from here.

## Three things to know before you trust it

**Some people have more than one login.** Roles and tokens hang off a login,
not a person, so turning off one login can leave another one working. Where
that applies, the tool will not let you remove the role and sends you to the
person's System tab instead, which is the screen that handles it properly.
Better than a button that looks like it worked.

**Adding a role can take something away.** When someone holds several roles,
TouchPoint does not add up what they allow. It picks one role and uses only
that one. So giving somebody an extra role can quietly turn off something they
could see yesterday. The By role tab shows where each role sits in that order.

**A blank is not proof.** If a role shows nothing, it means nothing turned up
where the tool looked, not that the role is unused. Same with the accounts that
have no birthday on them: most are kiosks and department mailboxes, but one is
a real person. Check before you act on either.

---

Free, like the rest of them. Questions or problems, open an issue.

<sub>Developer notes: `_rolemap/` holds the scanner that builds the role map
from TouchPoint's source. It is not uploaded and not needed on the server. Its
output is gitignored because this repo is public.</sub>
