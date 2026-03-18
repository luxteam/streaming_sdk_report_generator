from enum import Enum
from dataclasses import dataclass

TEMPLATE_PATH = "./template/"
WORKING_DIR_PATH = "./tmp_template/"
REPORT_FILE_PATH = "./report.docx"


class Jobs(Enum):
    Full_Samples = 1
    Full_Samples_Android = 2
    Win_OS = 3
    Win_Full = 4
    Win_APU = 5
    Win_Latency = 6
    Android_OS = 7
    Android_Full = 8
    Ubuntu_OS = 9
    Ubuntu_Full = 10
    # AMD_Full = 6
    Win_Long_Term = 11
    Android_Xiaomi_TV = 12
    Android_Chromecast_TV = 13


class Reports(Enum):
    summary = 1
    PUBG = 2
    Dota2_DX11 = 3
    Dota2_Vulkan = 4
    LoL = 5
    Heaven_Benchmark_DX9 = 6
    Valley_Benchmark_DX9 = 7
    Heaven_Benchmark_DX11 = 8
    Valley_Benchmark_DX11 = 9
    Heaven_Benchmark_OpenGL = 10
    Valley_Benchmark_OpenGL = 11


@dataclass
class Issue:
    key: str
    summary: str
    created_at: str
    severity: str
    url: str
