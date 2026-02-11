from __future__ import annotations

import json
import os
from dataclasses import dataclass


@dataclass
class DbConfig:
    host: str = "localhost"
    port: int = 5432
    dbname: str = "aman_inventory"
    user: str = "postgres"
    password: str = "postgres"


def load_db_config() -> DbConfig:
    env_url = os.getenv("AMAN_DB_URL")
    if env_url:
        # Use a DSN string directly if provided.
        return DbConfig(host=env_url, port=0, dbname="", user="", password="")

    path = os.path.join(os.path.dirname(__file__), "db_config.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return DbConfig(
            host=data.get("host", "localhost"),
            port=int(data.get("port", 5432)),
            dbname=data.get("dbname", "aman_inventory"),
            user=data.get("user", "postgres"),
            password=data.get("password", "postgres"),
        )
    return DbConfig()
