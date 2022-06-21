import json
import os
from pathlib import Path

from pydantic import BaseSettings


class Settings(BaseSettings):
    title_row_index: int = 1
    values_row_start_index: int = 2
    unique_cols: list[str] = ['F', 'L']
    sum_cols: list[str] = ['M', 'O']

    class Config:
        env_file = ".env"


def chek_default_config():
    env = Path('.env')
    if env.is_file():
        return

    settings = Settings().dict()

    with open('.env', 'w') as file:
        file.write('\n'.join(f'{key}={json.dumps(value)}' for key, value in settings.items()))
