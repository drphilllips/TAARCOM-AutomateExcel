from enum import Enum


class Company(Enum):
    DGK = "Digi-Key"
    MOU = "Mouser"
    ABR = "Abracon"
    NA = "Not Assigned"


class Flag(Enum):
    OOT = "Out of territory"  # Zip code not found
    CNF = "New account"  # Customer not found
    CNP = "Individual"  # Individual
    PNA = "No proper name associated"  # In customer-proper name map, but no associated proper name
    PNF = "Proper name not found"  # Proper name in customer-proper name map but not in Master Account List


class AbraconFlags(Enum):
    YELLOW_DESC = "NPI"  # new product introduction
    GREEN_DESC = "Large customer"
    ORANGE_DESC = "Alt brands"  # companies they bought out
    YELLOW_HEX = "FFFF00"
    GREEN_HEX = "92D050"
    ORANGE_HEX = "FFC000"


