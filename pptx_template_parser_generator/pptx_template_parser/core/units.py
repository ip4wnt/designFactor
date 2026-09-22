EMU_PER_CM = 360000


def emu_to_cm(value: int) -> float:
    return round(value / EMU_PER_CM, 2)


def pt_to_cm(value: float) -> float:
    return round(value * 0.0352778, 2)