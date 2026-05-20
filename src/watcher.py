"""
watch_and_refresh_pbi.py
────────────────────────
Surveille le dossier CSV partagé et déclenche un refresh PowerBI Desktop
uniquement quand les fichiers ont changé.

Usage :
    python watch_and_refresh_pbi.py

Lancement automatique au démarrage Windows → voir section "Installation"
en bas de ce fichier.

Dépendances :
    pip install watchdog pywin32
"""

import os
import sys
import time
import logging
import hashlib
import subprocess
import json
from pathlib import Path
from datetime import datetime

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
WATCHED_FOLDER   = Path(r'C:\\Users\\Rodri\\Desktop\\UTT\\P26\\crunch\\DATA\\Result')   # Dossier CSV
PBIX_FILE        = Path(r'C:\\Users\\Rodri\\Desktop\\UTT\\P26\\crunch\\dashboard\\Crunch_302B.pbix')  # Fichier PowerBI
PBI_EXE          = Path(r'C:\\Program Files\\Microsoft Power BI Desktop\bin\PBIDesktop.exe')
LOG_FILE         = Path(r'C:\\Users\\Rodri\\Desktop\\UTT\\P26\\crunch\\logs\watcher_log.txt')
STATE_FILE       = Path(r'C:\\Users\\Rodri\\Desktop\\UTT\\P26\\crunch\\logs\\file_state.json')   # Hash des fichiers précédents
CSV_PREFIX       = 'CRUNCH_'                                   # Surveiller uniquement ces fichiers
REFRESH_COOLDOWN = 3600   # Secondes d'attente minimum entre deux refreshs (évite les doubles)
CHECK_INTERVAL   = 3600   # Secondes entre deux vérifications de l'état des fichiers
# ─────────────────────────────────────────────────────────────────────────────

# Configuration du logger
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger(__name__)

last_refresh_time = 0


def get_file_hash(path: Path) -> str:
    """Calcule le hash MD5 d'un fichier pour détecter les changements."""
    try:
        h = hashlib.md5()
        with open(path, 'rb') as f:
            while chunk := f.read(8192):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ''


def load_state() -> dict:
    """Charge l'état précédent des fichiers (hash + timestamp)."""
    try:
        if STATE_FILE.exists():
            return json.loads(STATE_FILE.read_text(encoding='utf-8'))
    except Exception:
        pass
    return {}


def save_state(state: dict):
    """Sauvegarde l'état actuel des fichiers."""
    try:
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATE_FILE.write_text(json.dumps(state, indent=2), encoding='utf-8')
    except Exception as e:
        log.warning(f"Impossible de sauvegarder l'état : {e}")


def get_current_state() -> dict:
    """Calcule l'état actuel de tous les CSV surveillés."""
    state = {}
    if not WATCHED_FOLDER.exists():
        return state
    for f in WATCHED_FOLDER.glob(f'{CSV_PREFIX}*.csv'):
        state[f.name] = {
            'hash'     : get_file_hash(f),
            'size'     : f.stat().st_size,
            'modified' : f.stat().st_mtime,
        }
    return state


def files_have_changed(old_state: dict, new_state: dict) -> list:
    """
    Retourne la liste des fichiers qui ont changé.
    Détecte : nouveaux fichiers, fichiers modifiés, fichiers supprimés.
    """
    changed = []

    # Nouveaux fichiers ou fichiers modifiés
    for name, info in new_state.items():
        if name not in old_state:
            changed.append(f"{name} (nouveau)")
        elif info['hash'] != old_state[name]['hash']:
            changed.append(f"{name} (modifié)")

    # Fichiers supprimés
    for name in old_state:
        if name not in new_state:
            changed.append(f"{name} (supprimé)")

    return changed


def is_pbi_running() -> bool:
    """Vérifie si PowerBI Desktop est déjà ouvert."""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq PBIDesktop.exe', '/NH'],
            capture_output=True, text=True
        )
        return 'PBIDesktop.exe' in result.stdout
    except Exception:
        return False


def refresh_powerbi():
    """
    Déclenche l'actualisation de PowerBI Desktop.

    Stratégie :
    - Si PBI est déjà ouvert → envoie Ctrl+Alt+F5 (raccourci refresh universel)
    - Si PBI est fermé → ouvre le fichier .pbix (PBI rafraîchit au démarrage
      si l'option est activée dans le rapport)
    """
    global last_refresh_time

    now = time.time()
    if now - last_refresh_time < REFRESH_COOLDOWN:
        remaining = int(REFRESH_COOLDOWN - (now - last_refresh_time))
        log.info(f"Cooldown actif — prochain refresh possible dans {remaining}s")
        return

    try:
        if is_pbi_running():
            log.info("PowerBI Desktop détecté — envoi du raccourci refresh (Ctrl+Alt+F5)...")
            _send_pbi_refresh_hotkey()
        else:
            log.info(f"Ouverture de {PBIX_FILE.name}...")
            if PBIX_FILE.exists() and PBI_EXE.exists():
                subprocess.Popen([str(PBI_EXE), str(PBIX_FILE)])
            else:
                log.warning("Fichier .pbix ou PBIDesktop.exe introuvable — refresh ignoré.")
                return

        last_refresh_time = now
        log.info("Refresh PowerBI déclenché avec succès.")

    except Exception as e:
        log.error(f"Erreur lors du refresh PowerBI : {e}")


def _send_pbi_refresh_hotkey():
    """
    Envoie le raccourci clavier Ctrl+Alt+F5 à PowerBI Desktop via pywin32.
    Fonctionne même si la fenêtre n'est pas au premier plan.
    """
    try:
        import win32gui
        import win32con
        import win32api

        # Trouver la fenêtre PowerBI
        hwnd = None
        def enum_callback(h, _):
            nonlocal hwnd
            title = win32gui.GetWindowText(h)
            if 'Power BI' in title:
                hwnd = h
        win32gui.EnumWindows(enum_callback, None)

        if hwnd:
            # Mettre la fenêtre au premier plan brièvement
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            # Envoyer Ctrl+Alt+F5
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(win32con.VK_MENU,    0, 0, 0)
            win32api.keybd_event(win32con.VK_F5,      0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(win32con.VK_F5,      0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_MENU,    0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        else:
            log.warning("Fenêtre PowerBI non trouvée — ouverture du fichier à la place.")
            if PBIX_FILE.exists() and PBI_EXE.exists():
                subprocess.Popen([str(PBI_EXE), str(PBIX_FILE)])

    except ImportError:
        log.warning("pywin32 non installé — utilisez : pip install pywin32")
        log.warning("Tentative d'ouverture du fichier .pbix à la place...")
        if PBIX_FILE.exists() and PBI_EXE.exists():
            subprocess.Popen([str(PBI_EXE), str(PBIX_FILE)])


def main():
    log.info("=" * 60)
    log.info("Démarrage du watcher CRUNCH → PowerBI")
    log.info(f"Dossier surveillé : {WATCHED_FOLDER}")
    log.info(f"Fichier PowerBI   : {PBIX_FILE}")
    log.info(f"Intervalle        : {CHECK_INTERVAL}s | Cooldown : {REFRESH_COOLDOWN}s")
    log.info("=" * 60)

    if not WATCHED_FOLDER.exists():
        log.error(f"Dossier introuvable : {WATCHED_FOLDER}")
        log.error("Vérifiez la variable WATCHED_FOLDER dans la configuration.")
        sys.exit(1)

    # Charger l'état précédent
    previous_state = load_state()
    if previous_state:
        log.info(f"État précédent chargé : {len(previous_state)} fichier(s) connus")
    else:
        log.info("Premier démarrage — initialisation de l'état de référence")
        previous_state = get_current_state()
        save_state(previous_state)

    log.info("Surveillance active — en attente de changements...\n")

    try:
        while True:
            time.sleep(CHECK_INTERVAL)

            current_state = get_current_state()
            changed_files = files_have_changed(previous_state, current_state)

            if changed_files:
                log.info(f"Changements détectés ({len(changed_files)} fichier(s)) :")
                for f in changed_files:
                    log.info(f"  • {f}")

                # Sauvegarder le nouvel état
                save_state(current_state)
                previous_state = current_state

                # Déclencher le refresh PowerBI
                refresh_powerbi()
            else:
                log.debug("Aucun changement détecté.")

    except KeyboardInterrupt:
        log.info("Arrêt du watcher (Ctrl+C)")
    except Exception as e:
        log.error(f"Erreur inattendue : {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()


# ══════════════════════════════════════════════════════════════════════════════
# INSTALLATION — lancement automatique au démarrage Windows
# ══════════════════════════════════════════════════════════════════════════════
#
# Option A — Tâche planifiée (recommandé, fonctionne sans session ouverte) :
#
#   Dans PowerShell (admin) :
#
#   $action  = New-ScheduledTaskAction `
#       -Execute "python" `
#       -Argument "C:\CRUNCH\watch_and_refresh_pbi.py" `
#       -WorkingDirectory "C:\CRUNCH"
#
#   $trigger = New-ScheduledTaskTrigger -AtStartup
#
#   $settings = New-ScheduledTaskSettingsSet `
#       -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
#       -RestartCount 3 `
#       -RestartInterval (New-TimeSpan -Minutes 1)
#
#   Register-ScheduledTask `
#       -TaskName "CRUNCH PowerBI Watcher" `
#       -Action $action `
#       -Trigger $trigger `
#       -Settings $settings `
#       -RunLevel Highest
#
# Option B — Dossier Démarrage (plus simple, nécessite une session ouverte) :
#   Créez un raccourci vers ce script dans :
#   C:\Users\VotreNom\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
#
# ══════════════════════════════════════════════════════════════════════════════
# DÉPENDANCES
# ══════════════════════════════════════════════════════════════════════════════
#
#   pip install watchdog pywin32
#
# ══════════════════════════════════════════════════════════════════════════════
# OPTION AVANCÉE — PowerBI Service avec webhook (si vous avez une licence Pro)
# ══════════════════════════════════════════════════════════════════════════════
#
# Si vous publiez votre .pbix sur PowerBI Service, vous pouvez déclencher
# un refresh via l'API REST PowerBI au lieu du raccourci clavier.
# Remplacez refresh_powerbi() par :
#
#   import requests
#   TENANT_ID    = "votre-tenant-id"
#   CLIENT_ID    = "votre-client-id"
#   CLIENT_SECRET= "votre-secret"
#   DATASET_ID   = "votre-dataset-id"
#   GROUP_ID     = "votre-workspace-id"
#
#   def refresh_powerbi_service():
#       token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/token"
#       token = requests.post(token_url, data={
#           'grant_type': 'client_credentials',
#           'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
#           'resource': 'https://analysis.windows.net/powerbi/api'
#       }).json()['access_token']
#
#       requests.post(
#           f"https://api.powerbi.com/v1.0/myorg/groups/{GROUP_ID}/datasets/{DATASET_ID}/refreshes",
#           headers={'Authorization': f'Bearer {token}'}
#       )
#       log.info("Refresh PowerBI Service déclenché via API REST.")
#
# ══════════════════════════════════════════════════════════════════════════════