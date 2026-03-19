\# CASB Automation Framework



Automated CASB block verification framework for Versa SASE.

Tests that CASB Apps  are correctly blocked by CASB policy.



Built with: Playwright · Paramiko · pywinauto · Flask



\---



\## Folder Structure



```

casb\_new\_structure/

├── run.py                          ← Entry point. Add 1 line here per new app.

├── config.py                       ← All credentials, SSH settings, VOS config.

├── debug\_casb\_block\_alert\_popup\_finder.py  ← Debug tool for popup detection.

├── core/

│   ├── base\_activity.py            ← Base class for all app activities.

│   ├── browser\_handler.py          ← Browser helpers (tabs, screenshots, HAR).

│   ├── runner.py                   ← Generic test orchestrator.

│   ├── versa\_handler.py            ← CASB popup + SSH fast.log capture.

│   ├── vos\_info\_dump.py            ← VOS SSH commands (clear, dump, qosmos).

│   ├── decryption\_check.py         ← TLS inspection verification.

│   └── report\_generator.py         ← JSON + HTML report generation.

└── apps/

&#x20;   └── ms\_teams/

&#x20;       ├── app.yaml                ← App definition (keywords, activities).

&#x20;       ├── activities.py           ← MS Teams UI automation (TC1-TC4).

&#x20;       └── login\_handler.py        ← MS Teams login flow.

```



\---



\## Setup \& Installation



```bash

pip install playwright paramiko pywinauto flask waitress pyyaml requests cryptography

playwright install chromium

```



\---



\## How to Run



```bash

python run.py

&#x20; --applications "MS\_Teams"

&#x20; --account\_type personal

&#x20; --host 172.20.4.5

&#x20; --pwd versa123

&#x20; --ssh-user admin

&#x20; --org "ENDTOEND-Tenant-2"

&#x20; --access-policy "Default-Policy"

&#x20; --decrypt-policy "Default-Policy"

&#x20; --decrypt-rule "decryption\_rule\_casb"

&#x20; --decrypt-profile "decrypt\_profile"

&#x20; --casb-profile "casb\_mobile\_test\_rule"

&#x20; --casb-access-policy-rule "mobile\_test\_rule"

&#x20; --report-dir "C:\\Users\\admin\\Downloads\\CASB\_Reports"

&#x20; --qosmos True

&#x20; --send-email "user1@versa-networks.com,user2@versa-networks.com"

&#x20; --activities "post"

&#x20; --server-url "http://10.196.3.26:4012"

```



\### Arguments



| Argument | Purpose | Example |

|---|---|---|

| `--applications` | Comma-separated apps | `MS\_Teams` / `MS\_Teams,Instagram` |

| `--host` | VOS branch IP | `172.20.4.5` |

| `--pwd` | SSH password | `versa123` |

| `--ssh-user` | SSH username | `admin` |

| `--org` | VOS org name | `ENDTOEND-Tenant-2` |

| `--account\_type` | Account type(s) | `personal` / `corporate` / `personal,corporate` |

| `--activities` | TCs to run | see below |

| `--qosmos` | appid report\_metadata | `True` / `False` |

| `--report-dir` | Report output folder | `C:\\Users\\admin\\Downloads\\CASB\_Reports` |

| `--server-url` | CASB Results dashboard URL | `http://10.196.3.26:4012` |

| `--send-email` | Email report recipients | `a@versa.com,b@versa.com` |

| `--access-policy` | VOS access policy name | `Default-Policy` |

| `--decrypt-policy` | VOS decryption policy name | `Default-Policy` |

| `--decrypt-rule` | VOS decryption rule name | `decryption\_rule\_casb` |

| `--decrypt-profile` | VOS decryption profile name | `decrypt\_profile` |

| `--casb-profile` | VOS CASB profile name | `casb\_mobile\_test\_rule` |

| `--casb-access-policy-rule` | VOS CASB rule name | `mobile\_test\_rule` |



\### --activities combos



```

all                      -> run all activities and all TCs

post                     -> run all post TCs (TC1, TC2, TC3, TC4)

post\[1]                  -> run TC1 only

post\[2]                  -> run TC2 only

post\[3]                  -> run TC3 only

post\[4]                  -> run TC4 only

post\[1,2]                -> run TC1 and TC2

post\[1,2,3]              -> run TC1, TC2 and TC3

post\[1,2,3,4]            -> run all 4 TCs explicitly

post\[2,4]                -> run TC2 and TC4 (any combo works)

"post\[1,3] share\[1,4]"  -> run post TC1,TC3 AND share TC1,TC4

"post share"             -> run all TCs for both post and share

```



\---



\## Adding a New App



Only 3 things needed — core framework is never touched.



\*\*Step 1 — Create `apps/{app\_id}/app.yaml`\*\*

```yaml

name: Instagram

app\_id: instagram

app\_url: https://www.instagram.com

log\_match:

&#x20; keywords: \[instagram, post, app-activity for casb]

expected:

&#x20; application: instagram

&#x20; activity: post

&#x20; blocked\_by: casb

activities:

&#x20; post:

&#x20;   tc\_label: TC1

&#x20;   category: post

&#x20;   nav: "Home -> New Post -> Share"

```



\*\*Step 2 — Create `apps/{app\_id}/activities.py`\*\*

```python

from core.base\_activity import BaseActivity



class InstagramActivity(BaseActivity):

&#x20;   def \_open\_fresh\_tab(self): ...

&#x20;   def \_wait\_for\_app(self, page): ...

&#x20;   def \_do\_post(self, page, result, \*\*kwargs):

&#x20;       vsmd, har = self.\_before\_send(page, "TC1")

&#x20;       # ... UI clicks only ...

&#x20;       self.\_after\_send(page, result, vsmd, har, "TC1", None)

```



\*\*Step 3 — Add 1 line to `run.py`\*\*

```python

\_APP\_MAP = {

&#x20;   "ms\_teams" : \["personal", "corporate"],

&#x20;   "instagram": \["any"],   # <- this line

}

```



Then run:

```bash

python run.py --applications "Instagram" --host 172.20.4.5 --pwd versa123 --ssh-user admin

```



\---



\## Debug — CASB Popup Window Finder



When setting up a \*\*new app\*\* or on a \*\*new machine\*\*, run this tool first to identify

the exact Versa CASB AlertWindow title and class name.



```bash

python debug\_casb\_block\_alert\_popup\_finder.py

```



\*\*Steps:\*\*

1\. Run the script — it captures a baseline of all open windows

2\. Go to your app and manually perform the activity that CASB should block

3\. Script detects and prints any new windows that appear

4\. Look for the window marked `<- CASB POPUP (use this)` in the output



\*\*Example output:\*\*

```

\[15s] \*\*\* NEW WINDOW(S) DETECTED \*\*\*

&#x20;  TITLE   : 'AlertWindow'  <- CASB POPUP (use this)

&#x20;  CLASS   : 'HwndWrapper\[VersaSecureAccessClient.Alerts.exe;;...]'

&#x20;  BACKEND : win32



&#x20;  TITLE   : 'MediaContextNotificationWindow'  <- noise (ignore)

&#x20;  TITLE   : 'SystemResourceNotifyWindow'  <- noise (ignore)

```



Works for \*\*any app, any activity\*\* — you trigger the block manually.



\---



\## Results Dashboard



Accessible at: \*\*http://10.196.3.26:4012\*\*



\- View all runs with Pass/Fail, TC count, Trigger %, Sig IDs

\- Per-TC breakdown — CASB Block, Not Delivered, fast.log, fail reasons

\- Download ZIP, view HTML report, browse VOS dumps, HAR files, screenshots

\- Auto-upload via `--server-url` flag — no manual steps needed



\---



\## Git Workflow



\### Branches

```

main              <- stable, tested code only

├── amruta/apps   <- Amruta's work

├── lisari/apps   <- Lisari's work

├── lankesh/apps  <- Lankesh's work

└── hrutuja/apps  <- Hrutuja's work

```



\### First time setup

```bash

git clone https://github.com/TestCasbDoc/casb-automation.git

cd casb-automation

git checkout lankesh/apps    # use your own branch name

```



\### Daily workflow

```bash

\# Before starting — get latest

git pull



\# After making changes — save to GitHub

git add .

git commit -m "describe what you changed"

git push

```



\### Rules

\- Always `git pull` before starting work

\- Only work on your own branch

\- Never push directly to `main`

\- When your app is ready -> raise a Pull Request to merge into `main`

