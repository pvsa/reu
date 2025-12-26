import sys
import os
import subprocess

# === Modulprüfung ===
REQUIRED_MODULES = ["ldap3", "configparser"]

def check_modules():
    missing = []
    for module in REQUIRED_MODULES:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)

    if missing:
        print("Fehlende Python-Module erkannt:")
        for m in missing:
            print(f"  - {m}")
        print("\nBitte installiere sie mit, z. B.:")
        print(f"pip install {' '.join(missing)}\n")
        sys.exit(1)

# Modulprüfung sofort ausführen
check_modules()

# Jetzt Import nach erfolgreicher Prüfung:
import configparser
from ldap3 import Server, Connection, ALL, ALL_ATTRIBUTES, SUBTREE, core

def load_config(conf_path="conf/mbx-report.conf"):
    """Lädt die LDAP-Konfiguration aus einer externen Datei."""
    config = configparser.RawConfigParser() 
    if not os.path.exists(conf_path):
        raise FileNotFoundError(f"Konfigurationsdatei nicht gefunden: {conf_path}")
    config.read(conf_path)
    ldap_conf = config["ldap"]
    return {
        "server": ldap_conf.get("server"),
        "user": ldap_conf.get("user"),
        "password": ldap_conf.get("password"),
        "tree": ldap_conf.get("tree"),
    }

def main():
    try:
        conf = load_config()
        ldap_server = conf["server"]
        ldap_user = conf["user"]
        ldap_pass = conf["password"]
        ldap_tree = conf["tree"]

        # Verbindung herstellen
        server = Server(ldap_server, get_info=ALL)
        conn = Connection(server, user=ldap_user, password=ldap_pass, auto_bind=True)
        print("LDAP bind successful.\n")

        # Suche nach Organisationseinträgen
        conn.search(ldap_tree, "(objectclass=organization)", search_scope=SUBTREE, attributes=ALL_ATTRIBUTES)
        data = conn.entries

        kdn_dom = {}
        kdnlist = []

        for entry in data:
            attrs = entry.entry_attributes_as_dict
            if "o" in attrs and "destinationIndicator" in attrs:
                kunde = attrs["destinationIndicator"][0]
                kdn_dom.setdefault(kunde, []).append(attrs["o"][0])
                if kunde not in kdnlist:
                    kdnlist.append(kunde)

        kdnlist = list(set(kdnlist))
        total_mailboxes = 0

        for kd in kdnlist:
            print(f"{kd}")
            print("=" * 30)
            kcount = 0
            for domain in kdn_dom[kd]:
                print(f"Domain: {domain}")
                conn.search(ldap_tree, f"(o={domain})", search_scope=SUBTREE, attributes=ALL_ATTRIBUTES)
                entries2 = conn.entries
                j = 1
                for usrdn in entries2:
                    attrs2 = usrdn.entry_attributes_as_dict
                    if "mail" in attrs2:
                        mail = attrs2["mail"][0]
                        mailuser, _ = mail.split("@", 1)
                        maildir = f"/var/mail/vhosts/{domain}/{mailuser}"
                        if os.path.exists(maildir):
                            try:
                                result = subprocess.check_output(["du", "-sh", maildir], universal_newlines=True)
                                mbx_sz = result.split("\t")[0]
                            except subprocess.CalledProcessError:
                                mbx_sz = "0B"
                        else:
                            mbx_sz = "0B"
                        print(f"{j}. {mail}   {mbx_sz}")
                        j += 1
                print(f" Anzahl Mbx in Domain: {j-1}\n")
                kcount += (j - 1)
            print(f"Anzahl Mbx für Kunde: {kcount}\n\n")
            total_mailboxes += kcount

        print(f"Anzahl aller Mailboxen: {total_mailboxes}")
        conn.unbind()

    except core.exceptions.LDAPException as e:
        print(f"LDAP error: {e}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
