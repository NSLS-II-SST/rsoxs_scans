#imports
import datetime
import numpy as np
import matplotlib.pyplot as plt
import math
from copy import deepcopy
import warnings
from .defaults import *

def get_nexafs_scan_params(edge, speed = default_speed, ratios = None, quiet=False,**kwargs):
    """ 
    creates fly NEXAFS scan parameters, given an edge (which includes thresholds for different speed regions) base speed (eV/sec) and speed ratios between the different regions 
    Args:
        edge: String or tuple of 2 or more numbers (the only required entry)
            if string, this should be a key in the lookup tables refering to an absorption edge energy
            if tuple, this is the thresholds of the different ranges, i.e. start of pre edge, start of main edge, start of post edge, end etc.
                the length of the tuple should correspond to 1- the speed ratios (or a default can be chosen if between 2 and 6 thresholds are given)
        speed: String or Number (normal if not defined)
            if string, this should be in the lookup table of standard speeds for scans
            if number, this is the base speed in eV/sec
        speed_ratios: string or tuple of numbers
            if string, this should be in the lookup table of standard intervals
            if a tuple, this must have one less element than the edge tuple (either explicitely entered or from the lookup table)
                the values are the ratio of energy steps between the different regions defined by edge
                
    """
    
    edge_names = {"c":"Carbon",
              "carbonk": "Carbon",
              "ck": "Carbon",
              "n": "Nitrogen",
              "nitrogenk": "Nitrogen",
              "nk": "Nitrogen",
              "f": "Fluorine",
              "fluorinek": "Fluorine",
              "fk": "Fluorine",
              "o": "Oxygen",
              "oxygenk": "Oxygen",
              "ok": "Oxygen",
              "ca": "Calcium",
              "calciumk": "Calcium",
              "cak": "Calcium",
             }


     # these define the default edges of the intervals
    edges = {"carbon":(250,282,297,350),
             "oxygen":(500,525,540,560),
             "fluorine":(650,680,700,740),
             "aluminium":(1500,1560,1580,1600),
             "nitrogen":(370,397,407,440),
             "zincl": (1000,1015,1035,1085),
             "sulfurl":(150,160,170,200),
             "calcium":(320,345,355,380)
            }
     # these are the default interval ratios for each section
    ratios_table = {
                     "default 4":(5,1,5),
                     "default 5":(5,1,2,5),
                     "default 6":(5,2,1,2,5),
                     "default 2":(1,),
                     "default 3":(5,1),
                    }

    speed_table = {"normal": 0.2,"": 0.2,
                   "quick":0.4,
                   "fast":0.5,
                   "very fast":1,
                   "slow":0.1,
                   "very slow":0.05,
                  }
    
    edge_input = edge
    singleinput = False
    if isinstance(edge,str):
        if edge.lower() in edge_names.keys():
            edge = edge_names[edge.lower()]
            edge_input = edge.lower()
        if edge.lower() in edges.keys():
            edge = edges[edge.lower()]
    if not isinstance(edge,(tuple,list)):
        raise TypeError(f"invalid edge {edge} - no key of that name was found")
    if isinstance(speed,str):
        if speed.lower() in speed_table.keys():
            speed = speed_table[speed.lower()]
    if not isinstance(speed, (float,int)):
        raise TypeError(f"NEXAFS scan speed {speed} was not found or is not a valid number")
    if ratios == None or ratios == '':
        if str(edge_input).lower() in ratios_table.keys():
            ratios = ratios_table[edge_input.lower()]
        elif f"default {len(edge)}" in ratios_table.keys():
            ratios = ratios_table[f"default {len(edge)}"]
        else:
            ratios = (1,)*(len(edge)-1)
    else:
        if not isinstance(ratios,(tuple,list)):
            ratios = ratios_table[ratios]
    if not isinstance(ratios,(tuple,list)):
        raise TypeError(f"invalid ratios {ratios}")
    if len(ratios) + 1 != len(edge):
        raise ValueError(f'got the wrong number of intervals {len(ratios)} expected {len(edge)-1}')
    steps = 0
    scan_params = []
    time = 0
    for i, ratio in enumerate(ratios):
        scan_params += [(edge[i],edge[i+1],float(ratio)*float(speed))]
        time +=abs(edge[i+1] - edge[i]) / (float(ratio)*float(speed))
    if not quiet:
        # ------- remove this for production, it's just for looking at the output conveniently during development
        print(scan_params)
        print(len(scan_params))
        print(f'time : {datetime.timedelta(seconds=time)}')
        # --------
    
    return scan_params, time


def get_energies(edge,frames = default_frames, ratios = None,quiet=False,**kwargs):
    """
    creates a usable list of energies, given an edge an estimated number of frames (energies) and intervals
    Args:
        edge: String or tuple of 2 or more numbers or number
            if string, this should be a key in the lookup tables refering to an absorption edge energy
            if tuple, this is the thresholds of the different ranges, i.e. start of pre edge, start of main edge, start of post edge, end etc.
                the length of the tuple should correspond to 1- the length of intervals (or a default can be chosen if between 2 and 6 thresholds are given)
            if number, just return a single element array of that number - this is a very silly way to make such an array
        frames: String or Number
            if string, this should be in the lookup table of standard lengths for scans
            if number, this is the aim for the number of exposures needed in the region
                the algorithm will almost always exceed this estimate slightly
                precise numbers of exposures will not be possible using this function at this time
        intervals: string or tuple of numbers
            if string, this should be in the lookup table of standard intervals
            if a tuple, this must have one less element than the edge tuple (either explicitely entered or from the lookup table)
                the values should not be thought of as energy steps, but the ratio of energy steps between the different regions defined by edge
                
    benefits of this algorithm are:
    1.) the number of frames is always known, at least approximately - this is the main cause of confusion with the existing system
    2.) flexibility is maintained - any energy scan can be defined here, and with effort, even very precise energy choices are possible with just three relatively plain text inputs
    3.) estimation of time should be considerably more accurate than looking up previous scans.  Here we can have a very simple algorithm to estimate the time of a scan
    4.) maintains very simple standard scans that will always be the same - if you don't change the inputs, and only choose an edge, or a standard frame the same energies will be taken each time
    downsides:
    1.) the exact energies taken might not be obvious.  all of the exact energies in the edge definition will always be taken, but the exact step between is not necessarily forseeable
    
    definite additions before put into practice
    1.) define all look up tables in persistant database dict.  allow new definitions.  defaults overwritten by these tables on BSUI load (through redis for now)
    2.) add sim mode, where no errors are thrown, but explainations / estimations are returned as a string - return number of exposures for time estimation
    3.) no plotting?  or maybe just as option when in sim mode
    
    possible additions:
    1.) exposure times - added this as seperate function below
    2.) direct time estimation output - added to function for exposure times below
    """
    
    edge_names = {"c":"Carbon",
              "carbonk": "Carbon",
              "ck": "Carbon",
              "n": "Nitrogen",
              "nitrogenk": "Nitrogen",
              "nk": "Nitrogen",
              "f": "Fluorine",
              "fluorinek": "Fluorine",
              "fk": "Fluorine",
              "o": "Oxygen",
              "oxygenk": "Oxygen",
              "ok": "Oxygen",
              "ca": "Calcium",
              "calciumk": "Calcium",
              "cak": "Calcium",
             }


     # these define the default edges of the intervals
    edges = {"carbon":(250,270,282,287,292,305,350),
             "oxygen":(510,525,540,560),
             "fluorine":(670,680,690,700,740),
             "aluminium":(1540,1560,1580,1600),
             "nitrogen":(380,397,407,440),
             "zincl": (1000,1015,1035,1085),
             "sulfurl":(150,160,170,200),
             "calcium":(320,340,345,349,349.5,352.5,353,355,360,380)
            }
     # these are the default interval ratios for each section
    intervals_table = {"carbon":(5,1,.1,.2,1,5),
                     "carbon nonaromatic":(5,1,.2,.1,1,5),
                     "default 4":(2,.2,2),
                     "default 5":(2,.2,.6,2),
                     "default 6":(2,.6,.2,.6,2),
                     "default 2":(2,),
                     "default 3":(2,.2),
                     "calcium": (5,1,0.5,0.1,0.25,0.1,0.5,1,5),
                    }

    frames_table = {"full": 112,"": 112,
              "short":56,
              "very short":40,
             }
    edge_input = edge
    singleinput = False
    if isinstance(edge,str):
        if edge.lower() in edge_names.keys():
            edge = edge_names[edge.lower()]
            edge_input = edge.lower()
        if edge.lower() in edges.keys():
            edge = edges[edge.lower()]
    if isinstance(edge_input,(float,int)):
        edge = (edge_input,edge_input)
        singleinput = True
    if not isinstance(edge,(tuple,list)):
        raise TypeError(f"invalid edge {edge} - no key of that name was found")
    if isinstance(frames,float):
        if math.isnan(frames):
            frames = 'full'
    if isinstance(frames,str):
        if frames.lower() in frames_table.keys():
            frames = frames_table[frames.lower()]
    if not isinstance(frames, (int,float)):
        raise TypeError(f"frame number {frames} was not found or is not a valid number")
    
    if ratios == None or ratios == "":
        if str(edge_input).lower() in intervals_table.keys():
            ratios = intervals_table[edge_input.lower()]
        elif f"default {len(edge)}" in intervals_table.keys():
            ratios = intervals_table[f"default {len(edge)}"]
        else:
            ratios = (1,)*(len(edge)-1)
    else:
        if not isinstance(ratios,(tuple,int,float,list)):
            ratios = intervals_table[ratios]
    if not isinstance(ratios,(tuple,int,float,list)):
        raise TypeError(f"invalid intervals {ratios}")
    if len(ratios) + 1 != len(edge):
        raise ValueError(f'got the wrong number of intervals {len(ratios)} expected {len(edge)-1}')
    steps = 0
    multiple = 1
    numpnts = np.zeros_like(ratios)
    for i, interval in enumerate(ratios):
        numpnts[i] = np.round(np.abs(edge[i+1] - edge[i])/(interval*multiple))
    steps =sum(numpnts)
    multiple  *= steps / frames # if there are too many steps, multiple will reduce to approximately match the frames needed
    steps = 0
    energies = np.zeros(0)
    for i, interval in enumerate(ratios):
        if interval < 0.01:
            raise ValueError(f'ratio value of {interval} invalid')
        if multiple < 0.01:
            raise ValueError(f'ratio value of {multiple} invalid')
        numpnt = max(1,int(np.round(np.abs(edge[i+1] - edge[i])/max(0.01,interval*multiple)))) # get the number of points using this multiple
        at_end = i==len(ratios)-1 and not singleinput # if we are at the end, add the last point (built into linspace)
        energies = np.append(energies,np.around(np.linspace(edge[i],edge[i+1],numpnt+at_end,endpoint=at_end)*2,1)/2) # rounds to nearest 0.05 eV for clarity
    if not quiet:
        # ------- remove this for production, it's just for looking at the output conveniently during development
        print(energies)
        print(len(energies))
        plt.plot(energies,marker='x',markersize=5,linewidth=0.1)
        plt.show()
        # --------
    return energies


def construct_exposure_times(energies,exposure_time=1,repeats=1,quiet=False):
    """
    construct an array of exposure times to go with the array of energies input
    also estimate the total time for the scan, using a fixed 
    
    inputs:
        energies: an array or list of energies (floats or ints)
        exposure:
            (number) set all exposure times to this default value
            (list) assume this is a default value and a list of logical tests followed by their respective exposure times
                example [1,('less_than',270),10,('between',285,288),0.1]  the only logical operators allowed are "less_than", "between", "equals", and "greater_than"
    
    outputs:
        times : an array the same length of energies with exposure times
        time : the seconds estimated for the scan
    """
    if 1 > int(repeats) or 100 < int(repeats):
        raise ValueError("repeats must be a positive integer between 0 and 100")
    if not isinstance(energies,np.ndarray):
        raise ValueError("Invalid list of energies")
    times = energies.copy()
    if exposure_time == '':
        exposure_time = 1
    if isinstance(exposure_time,(float,int)):
        times[:] = float(exposure_time)
    elif isinstance(exposure_time,list):
        times[:] = float(exposure_time[0])
        for test,value in zip(exposure_time[1::2], exposure_time[2::2]):
            if test[0] in ['less_than','less than'] :
                times[energies < test[1]] = value
                #print(f"testing if energies are less than {test[1]} and setting them to {value}")
            elif test[0] in ['greater_than','greater than']:
                times[energies > test[1]] = value
                #print(f"testing if energies are greater than {test[1]} and setting them to {value}")
            elif test[0] == 'between':
                times[(test[1] < energies) * (energies < test[2])] = value
                #print(f"testing if energies are between {test[1]} and {test[2]} and setting them to {value}")
            elif test[0] == 'equals':
                times[test[1] == energies] = value
                #print(f"testing if energies are equal to {test[1]} and setting them to {value}")
            else:
                raise ValueError(f"Invalid test, only less_than, greater_than, and between are accepted, got {test}")
    calc_times = times * repeats
    calc_times += 1*(repeats-1) # one second overhead between repeated exposures\
    time = sum(calc_times) + 4*len(times)
    if time>dafault_warning_step_time:
        warnings.warn('WARNING: step will take more than 30 minutes')
    return times, time
    