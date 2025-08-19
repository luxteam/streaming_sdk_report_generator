from enum import Enum
from dataclasses import dataclass

TEMPLATE_PATH = "./template/"
WORKING_DIR_PATH = "./tmp_template/"
REPORT_FILE_PATH = "./report.docx"


class Jobs(Enum):
    # Full_Samples = 1
    Win_Full = 1
    Win_APU = 2
    Win_Latency = 3
    Android_Full = 4
    Ubuntu_Full = 5
    AMD_Full = 6
    Win_Long_Term = 7
    Android_Xiaomi_TV = 8
    Android_Chromecast_TV = 9


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
