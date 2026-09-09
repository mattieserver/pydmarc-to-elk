from configparser import ConfigParser

CONFIG_PATH = "Settings/config.ini"

DEFAULTS = {
    "email": {
        "tenant_id": "00000000-0000-0000-0000-000000000000",
        "client_id": "00000000-0000-0000-0000-000000000000",
        "secret": "app_registration_client_secret",
        "mailbox_id": "mailauth-reports@example.com",
        "delete_processed": "no",
    },
    "elk": {
        "host": "192.168.0.1",
        "port": "9200",
        "mode": "read",  # set read or write
        "auth": "no",  # set yes or no
        "user": "username_elastic",
        "password": "password_elastic",
        "verify_certs": "yes",  # set no only for a self-signed cluster
    },
}

CONFIG = ConfigParser()
CONFIG.read(CONFIG_PATH)

for section, options in DEFAULTS.items():
    if not CONFIG.has_section(section):
        CONFIG.add_section(section)
    for option, value in options.items():
        if not CONFIG.has_option(section, option):
            CONFIG.set(section, option, value)

with open(CONFIG_PATH, "w") as f:
    CONFIG.write(f)

print("Wrote {}, missing options filled with defaults.".format(CONFIG_PATH))
