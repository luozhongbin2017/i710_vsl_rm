# attempt to use VISSIM COM interface through Pytyhon win32com
# ref : Mark Hammond, Andy Robinson (2000, Jan). Python Programming on Win32, [Online]. Available: http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/
# ref : VISSIM 5.30-04 - COM Interface Manual, PTV group, Karlsrueh, Germany, 2011.

import ctypes.wintypes
import email_interface
import inspect
import os
import pywintypes
import win32api
import win32com.client as com
import win_interface

def getProgramName():
    # ref : In Python, how do I get the path and name of the file that is currently executing?, 2008, [Online] Available at : http://stackoverflow.com/questions/50499/in-python-how-do-i-get-the-path-and-name-of-the-file-that-is-currently-executin
    PROGRAM_FULL_NAME = inspect.getfile(inspect.currentframe())
    PROGRAM_NAME = os.path.split(PROGRAM_FULL_NAME)[-1]
    return PROGRAM_NAME

def mkdir(path):
    ''' recursively create given path '''
    upper, name = os.path.split(path)
    if not os.path.exists(upper):
        mkdir(upper)
    # if path already exists, do not create. Otherwise, error will occur
    if not os.path.exists(path):
        os.mkdir(path)  


def calcRandomSeeds(nMonteCarlo):
    # list comprehension
    # same as {3i + 1 | 0 <= i < nMonteCarlo}
    return [3*i + 1 for i in range(nMonteCarlo)]

def writeHeader(fileName, title, Nsec, firstColumnHeader = "t/s", columnHeader = "Sec"):
    '''write header line of a data file
    Parameters
    ----------
    fileName : string
        file name
    title : string
        first line to be written
    Nsec : int
        number of columns
    firstColumnHeader : string
        header of the first column
    columnHeader : string
        header of each column. Repeated Nsec times with headers
        example : Nsec = 2, columnHeader = "Sec"
        -> "Sec1\tSec2\t"
    '''
    speedFile = open(fileName, 'w')
    # first line
    speedFile.write(title)
    speedFile.write('\n')
    # column headers
    #   first column header
    speedFile.write(firstColumnHeader)
    speedFile.write('\t')

    columnHeaderTemplate = columnHeader+"%d\t"
    # each column
    for i in xrange(Nsec):
        speedFile.write(columnHeaderTemplate % (i + 1))
    speedFile.write('\n')
    speedFile.close()

def runSimulation(simulationTime_sec, startTime_sec, endTime_sec, idxScenario, idxControler, idxLaneClosure, randonSeed, folderDir, networkDir):
    '''run PTV VISSIM simulation using given arguments
    Parameters
    ----------
        simulationTime_sec : int
            Simulation time (s)
        startTime_sec : int
            Time accident starts (s)
        endTime_sec : int
            Time accident ends (s)
        vehicleInput_veh_hr : int
            vehicle input (veh/hr)
        idxScenario : int
            Index of simulation scenarios
        idxController : int
            Controller Index
        idxLaneClosure : int
            Whether Lane Change control is added
        randomSeed : int
            seed for random numbers
        folderDir : string, path
            location of data files
        networkDir : string, path
            location of network file
    '''

    # Link Groups
    link_groups = [
                {'MAINLINE': (84,), 'ONRAMP': (), 'OFFRAMP': (), 'DC': 1, 'VSL': (1, 2, 3, 4)},
                {'MAINLINE': (83,), 'ONRAMP': (), 'OFFRAMP': (), 'DC': 2, 'VSL': (5, 6, 7, 8)},
                {'MAINLINE': (81,), 'ONRAMP': (82,), 'OFFRAMP': (80,), 'DC': 3, 'VSL': (9, 10, 11, 12, 13, 69) },
                {'MAINLINE': ()}
    ]




     [(84,),
                   (83,),
                   (81,),
                   (79,),
                   (77, 76),
                   (75, 74),
                   (72, 70),
                   (69,),
                   (68, 66,),
                   (65,),
                   (63, 62, 60,),
                   (59,),
                   (58,),
                   (57, 55, 51),
                   (50, 19, 18),
                   (17, 339),
                   (338,),
                   (337, 310, 311, 309, 307, 306, 304)
    ]

    #Data Collector Groups
    DC_groups = [()

    ]




    Nsec = 18   # number of sections
    simResolution = 5





    
