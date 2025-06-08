#!/usr/bin/env python3
# Kawin Viriyaprasopsook<kawin.v@kkumail.com>
# 2025-05-09
# sudo apt-get -y install pcscd python-pyscard python-pil

from dataclasses import dataclass
from pathlib import Path
from smartcard.System import readers
from smartcard.util import toHexString
import sys
from typing import Callable, List, Optional

# Decoder function
def thai2unicode(data: List[int]) -> str:
    return (
        bytes(data)
        .decode('tis-620', errors='replace')
        .replace('#', ' ')
        .strip()
    )

@dataclass(frozen=True)
class APDUCommand:
    ins: List[int]
    label: str
    decoder: Callable[[List[int]], str] = thai2unicode

class SmartCard:
    SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
    APPLET = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]

    def __init__(self, connection):
        self.conn = connection
        self.req: List[int] = []

    def connect(self):
        self.conn.connect()
        atr = self.conn.getATR()
        print("ATR:", toHexString(atr))
        self.req = [0x00, 0xC0, 0x00, 0x01] if atr[:2] == [0x3B, 0x67] else [0x00, 0xC0, 0x00, 0x00]

    def transmit(self, apdu: List[int]) -> (List[int], int, int):
        return self.conn.transmit(apdu)

    def initialize(self):
        sw1, sw2 = self.transmit(self.SELECT + self.APPLET)[1:]
        print(f"Select Applet: {sw1:02X} {sw2:02X}")

    def get_data(self, cmd: List[int]) -> List[int]:
        data, sw1, sw2 = self.transmit(cmd)
        data, sw1, sw2 = self.transmit(self.req + [cmd[-1]])
        return data

    def read_field(self, cmd: APDUCommand) -> str:
        data = self.get_data(cmd.ins)
        result = cmd.decoder(data)
        print(f"{cmd.label}: {result}")
        return result

    def read_photo(self, cid: str, segments: int = 20):
        base = [0x80, 0xB0, 0x00, 0x78, 0x00, 0x00, 0xFF]
        photo = bytearray()
        for i in range(1, segments + 1):
            cmd = base.copy()
            cmd[4] = i
            photo.extend(self.get_data(cmd))
        filename = Path(f"{cid}.jpg")
        filename.write_bytes(photo)
        print(f"Photo saved as {filename}")

def select_reader() -> Optional[object]:
    rlist = readers()
    if not rlist:
        print("No smartcard readers found.")
        return None
    print("Available readers:")
    for i, r in enumerate(rlist):
        print(f"  [{i}] {r}")
    try:
        choice = int(input("Select reader [0]: ") or 0)
    except ValueError:
        choice = 0
    return rlist[min(max(choice, 0), len(rlist) - 1)]

def main():
    reader = select_reader()
    if reader is None:
        sys.exit(1)

    conn = reader.createConnection()
    card = SmartCard(conn)
    card.connect()
    card.initialize()

    commands = [
        APDUCommand([0x80, 0xB0, 0x00, 0x04, 0x02, 0x00, 0x0D], "CID"),
        APDUCommand([0x80, 0xB0, 0x00, 0x11, 0x02, 0x00, 0x64], "TH Fullname"),
        APDUCommand([0x80, 0xB0, 0x00, 0x75, 0x02, 0x00, 0x64], "EN Fullname"),
        APDUCommand([0x80, 0xB0, 0x00, 0xD9, 0x02, 0x00, 0x08], "Date of birth"),
        APDUCommand([0x80, 0xB0, 0x00, 0xE1, 0x02, 0x00, 0x01], "Gender"),
        APDUCommand([0x80, 0xB0, 0x00, 0xF6, 0x02, 0x00, 0x64], "Card Issuer"),
        APDUCommand([0x80, 0xB0, 0x01, 0x67, 0x02, 0x00, 0x08], "Issue Date"),
        APDUCommand([0x80, 0xB0, 0x01, 0x6F, 0x02, 0x00, 0x08], "Expire Date"),
        APDUCommand([0x80, 0xB0, 0x15, 0x79, 0x02, 0x00, 0x64], "Address"),
    ]

    cid = ""
    for cmd in commands:
        result = card.read_field(cmd)
        if cmd.label == "CID":
            cid = result

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Error:", e)
    finally:
        sys.exit()