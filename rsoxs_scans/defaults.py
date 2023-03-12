"""Contains default measurement,timeout parameters as well as list of allowed actions.

"""

# set the defaults one place
default_frames = "full"
default_repeats = 1
default_speed = "normal"
default_cycles = 0
default_diameter = 1.8
default_spiral_step = 0.3
default_exposure_time = 1
default_warning_step_time = 1800
current_version = "2023-1 Version 1.0"

actions = {
    "load_configuration",  # high level load names RSoXS configuration
    "rsoxs_scan_core",  # high level run a single RSoXS plan
    "temp",  # change the temperature
    "spiral_scan_core",  # high spiral scan a sample
    "move",  # move an ophyd object
    "load_sample",  # high level load a sample
    "message",  # message the user - no action
    "diode_low",  # set diodes range setting to low
    "diode_high",  # set diode range setting to high
    "nexafs_scan_core",  # high level run a single NEXAFS scan
    "error",  # raise an error - should never get here.
}
"""
diode_low          : set()
spiral_scan_core   : {'energy', 'exposure', 'pol', 'stepsize', 'diameter', 'angle', 'grating', 'dets'}
temp               : {'wait', 'temp'}
load_configuration : {'configuration'}
move               : {'position', 'motor'}
diode_high         : set()
nexafs_scan_core   : {'cycles', 'pol', 'angle', 'energies', 'speeds', 'grating'}
message            : set()
load_sample        : {'sample'}
rsoxs_scan_core    : {'detnames', 'repeats', 'temperatures', 'temps_with_locations', 'polarizations', 'energies', 'grating', 'locations', 'times'}
"""


# look up table for aliases of edges
edge_names = {
    "c": "Carbon",
    "carbon": "Carbon",
    "carbonk": "Carbon",
    "ck": "Carbon",
    "n": "Nitrogen",
    "nitrogen": "Nitrogen",
    "nitrogenk": "Nitrogen",
    "nk": "Nitrogen",
    "f": "Fluorine",
    "fluorine": "Fluorine",
    "fluorinek": "Fluorine",
    "fk": "Fluorine",
    "o": "Oxygen",
    "oxygen": "Oxygen",
    "oxygenk": "Oxygen",
    "ok": "Oxygen",
    "ca": "Calcium",
    "calcium": "Calcium",
    "calciumk": "Calcium",
    "cak": "Calcium",
}


# these define the default edges of the intervals for rsoxs scans
rsoxs_edges = {
    "carbon": (250, 270, 282, 287, 292, 305, 350),
    "oxygen": (510, 525, 540, 560),
    "fluorine": (670, 680, 690, 700, 740),
    "aluminium": (1540, 1560, 1580, 1600),
    "nitrogen": (380, 397, 407, 440),
    "zincl": (1000, 1015, 1035, 1085),
    "sulfurl": (150, 160, 170, 200),
    "calcium": (320, 340, 345, 349, 349.5, 352.5, 353, 355, 360, 380),
}

# these are the default interval ratios for each section for rsoxs scans
rsoxs_ratios_table = {
    "carbon": (5, 1, 0.1, 0.2, 1, 5),
    "carbon nonaromatic": (5, 1, 0.2, 0.1, 1, 5),
    "default 4": (2, 0.2, 2),
    "default 5": (2, 0.2, 0.6, 2),
    "default 6": (2, 0.6, 0.2, 0.6, 2),
    "default 2": (2,),
    "default 3": (2, 0.2),
    "calcium": (5, 1, 0.5, 0.1, 0.25, 0.1, 0.5, 1, 5),
}

# aliases for default numbers of frames in rsoxs scans
frames_table = {
    "full": 112,
    "": 112,
    "short": 56,
    "very short": 40,
}

# these define the default edges of the intervals for nexafs scans
nexafs_edges = {
    "carbon": (250, 282, 297, 350),
    "oxygen": (500, 525, 540, 560),
    "fluorine": (650, 680, 700, 740),
    "aluminium": (1500, 1560, 1580, 1600),
    "nitrogen": (370, 397, 407, 440),
    "zincl": (1000, 1015, 1035, 1085),
    "sulfurl": (150, 160, 170, 200),
    "calcium": (320, 345, 355, 380),
}

# these are the default speed ratios for each section in nexafs scans
nexafs_ratios_table = {
    "default 4": (5, 1, 5),
    "default 5": (5, 1, 2, 5),
    "default 6": (5, 2, 1, 2, 5),
    "default 2": (1,),
    "default 3": (5, 1),
}

# these are the default speed aliases for NEXAFS scans
nexafs_speed_table = {
    "normal": 0.2,
    "": 0.2,
    "quick": 0.4,
    "fast": 0.5,
    "very fast": 1,
    "slow": 0.1,
    "very slow": 0.05,
}

# List of Valid Measurement Configurations
config_list = [
    "WAXSNEXAFS",
    "WAXS",
    "SAXS",
    "SAXSNEXAFS",
    "SAXS_liquid",
    "WAXS_liquid",
]

rsoxs_configurations = [
    "WAXS",
    "SAXS",
    "SAXS_liquid",
    "WAXS_liquid",
]

empty_acq = {
    "sample_id": "",
    "configuration": "",
    "type": "",
    "priority": "",
    "edge": "",
    "ratios": "",
    "frames": "",
    "repeats": "",
    "speed": "",
    "cycles": "",
    "diameter": "",
    "spiral_step": "",
    "pol_mode": "",
    "polarizations": "",
    "angles": "",
    "exposure_time": "",
    "grating": "",
    "diode_range": "",
    "group": "",
    "slack_message_start": "",
    "slack_message_end": "",
    "acquisition_notes": "",
    "uid": "",
    "runs": "",
}

empty_sample = {
    "bar_name": "",
    "sample_id": "",
    "sample_name": "",
    "project_name": "",
    "institution": "",
    "proposal_id": "",
    "bar_spot": "",
    "front": "",
    "grazing": "",
    "angle": "",
    "height": "",
    "sample_desc": "",
    "project_desc": "",
    "sample_priority": "",
    "notes": "",
    "sample_set": "",
    "components": "",
    "composition": "",
    "thickness": "",
    "density": "",
    "sample_date": "",
    "location": [],
    "bar_loc": {},
    "proposal": "",
    "SAF": "",
    "analysis_dir": "",
    "data_session": "",
    "acq_history": [],
}
