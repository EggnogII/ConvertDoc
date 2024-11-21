from enum import Enum

class horizontal_resolution(Enum):
    UHD = 3840
    QHD = 2560
    WUXGA = 1920
    FHD = 1920
    SXGA = 1280
    XGA = 1024
    SVGA = 800
    VGA = 640
    QVGA = 320

class vertical_resolution(Enum):
    UHD = 2160
    QHD = 1440
    WUXGA = 1200
    FHD = 1080
    SXGA = 1024
    XGA = 768
    SVGA = 600
    VGA = 480
    QVGA = 240