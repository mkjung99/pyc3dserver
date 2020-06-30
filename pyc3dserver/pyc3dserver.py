"""
MIT License

Copyright (c) 2020 Moon Ki Jung

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

__author__ = 'Moon Ki Jung, https://github.com/mkjung99/pyc3dserver'
__version__ = '0.0.7'

import os
import pythoncom
import win32com.client as win32
import traceback
import numpy as np
from scipy.interpolate import InterpolatedUnivariateSpline
import logging

logger_name = 'pyc3dserver'
logger = logging.getLogger(logger_name)
logger.setLevel('CRITICAL')
logger.addHandler(logging.NullHandler())

def init_logger(logger_lvl='WARNING', c_hdlr_lvl='WARNING', f_hdlr_lvl='ERROR', f_hdlr_f_mode='w', f_hdlr_f_path=None):
    """
    Initialize the logger of pyc3dserver module.

    Parameters
    ----------
    logger_lvl : str or int, optional
        Level of the logger itself. The default is 'WARNING'.
    c_hdlr_lvl : str or int, optional
        Level of the console handler in the logger. The default is 'WARNING'.
    f_hdlr_lvl : str or int, optional
        Level of the file handler in the logger. The default is 'ERROR'.
    f_hdlr_f_mode : str, optional
        File mode of the find handler in the logger. The default is 'w'.
    f_hdlr_f_path : str, optional
        File path of the file handler. The default is None.
        If this value is None, then there will be no file handler in the logger. 
        
    Returns
    -------
    logger : logging.Logger
        Logger object.

    """
    logger.setLevel(logger_lvl)
    while logger.hasHandlers():
        logger.removeHandler(logger.handlers[0])    
    if not logger.handlers:
        c_hdlr = logging.StreamHandler()
        c_hdlr.setLevel(c_hdlr_lvl)
        c_fmt = logging.Formatter('<%(name)s> - [%(levelname)s] - %(funcName)s() - %(message)s')
        c_hdlr.setFormatter(c_fmt)
        logger.addHandler(c_hdlr)
        if f_hdlr_f_path is not None:
            f_hdlr = logging.FileHandler(f_hdlr_f_path, mode=f_hdlr_f_mode)
            f_hdlr.setLevel(f_hdlr_lvl)
            f_fmt = logging.Formatter('%(asctime)s - <%(name)s> - [%(levelname)s] - %(funcName)s() - %(message)s')
            f_hdlr.setFormatter(f_fmt)
            logger.addHandler(f_hdlr)
    return logger

def reset_logger():
    """
    Reset the logger by setting its level as 'CRITICAL' and removing all its handlers.

    Returns
    -------
    None.

    """
    while logger.hasHandlers():
        logger.removeHandler(logger.handlers[0])    
    logger.setLevel('CRITICAL')       
    return None

def c3dserver(msg=True, log=False):
    """
    Initialize C3DServer COM interface using win32com.client.Dispatch().
    
    Also shows the relevant information of C3DServer status such as
    registration mode, version, user name and organization.    

    Parameters
    ----------
    msg : bool, optional
        Whether to show the information of C3Dserver. The default is True.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.

    """
    try:
        itf = win32.Dispatch('C3DServer.C3D')
        # itf = win32.dynamic.Dispatch('C3DServer.C3D')
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None        
    reg_mode = itf.GetRegistrationMode()
    ver = itf.GetVersion()
    user_name = itf.GetRegUserName()
    user_org = itf.GetRegUserOrganization()
    if msg:
        print('=============================================')
        if reg_mode == 0:
            print('Unregistered C3Dserver')
        elif reg_mode == 1:
            print('Evaluation C3Dserver')
        elif reg_mode == 2:
            print('Registered C3Dserver')
        print('Version: ', ver)
        print('User: ', user_name)
        print('Organization: ', user_org)
        print('=============================================')
    if log:
        if reg_mode == 0:
            logger.info('Unregistered C3Dserver')
        elif reg_mode == 1:
            logger.info('Evaluation C3Dserver')
        elif reg_mode == 2:
            logger.info('Registered C3Dserver')        
        logger.info(f'Version: {ver}')
        logger.info(f'User: {user_name}')
        logger.info(f'Organization: {user_org}')
    return itf

def open_c3d(itf, f_path, strict_param_check=False, log=False):
    """
    Open a C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    f_path : str
        Path of the input C3D file to open.
    strict_param_check: bool, optional
        Whether to enable strict parameter checking or not. The deafult is False.
    log: bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    if log: logger.debug(f'Opening the file: "{f_path}"')
    if not os.path.exists(f_path):
        if log: logger.error('File path does not exist!')
        return False
    try:
        ret = itf.Open(f_path, 3)
        if strict_param_check:
            itf.SetStrictParameterChecking(1)
        else:
            itf.SetStrictParameterChecking(0)        
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return False        
    if ret == 0:
        if log: logger.info(f'File is opened successfully.')
        return True
    else:
        if log: logger.info(f'File can not be opened.')
        return False

def save_c3d(itf, f_path='', f_type=-1, compress_param_blocks=False, log=False):
    """
    Save a C3D file.
    
    If 'f_path' is given an empty string, this function will overwrite the opened existing C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    f_path : str, optional
        Path of the output C3D file to save. The default is ''.
    f_type : int, optional
        Type of saving file. -1 means that the data is saved to the existing file type.
        1 for Intel(MS-DOS) format, 2 for DEC format, 3 for SGI format.
    compress_param_blocks: bool, optional
        Whether to remove any empty parameter blocks when the C3D file is saved. The default is False.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    bool
        True or False.

    """
    if log: logger.debug(f'Saving the file: "{f_path}"')
    try:
        if compress_param_blocks:
            itf.CompressParameterBlocks(1)
        else:
            itf.CompressParameterBlocks(0)
        ret = itf.SaveFile(f_path, f_type)
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return False
    if ret == 1:
        if log: logger.info(f'File is saved successfully.')
        return True
    else:
        if log: logger.info(f'File can not be saved.')
        return False
    return 

def close_c3d(itf, log=False):
    """
    Close a C3D file that has been previously opened and releases the memory.
    
    This function does not automatically save the C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    NoneType
        None.

    """
    if log: logger.info(f'File is closed.')
    return itf.Close()


def get_file_type(itf, log=False):
    """
    Return the file type of an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    str or None
        File type.

    """
    dict_f_type = {1:'INTEL', 2:'DEC', 3:'SGI'}
    try:
        f_type = itf.GetFileType()
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None      
    return dict_f_type.get(f_type, None)

def get_data_type(itf, log=False):
    """
    Return the data type of an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    str or None
        Data type.

    """
    dict_data_type = {1:'INTEGER', 2:'REAL'}
    try:
        data_type = itf.GetDataType()
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None    
    return dict_data_type.get(data_type, None)

def get_first_frame(itf, log=False):
    """
    Give you the first frame of video data from an open C3D file.
    
    This information is usually taken from the header record of the file.
    However, if the TRIAL:ACTUAL_START_FIELD and TRIAL:ACTUAL_END_FIELD parameters are present,
    the values from those parameters are used.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    int
        The first 3D frame number.

    """
    try:
        first_fr = itf.GetVideoFrame(0)
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None    
    return np.int32(first_fr)

def get_last_frame(itf, log=False):
    """
    Give you the last frame of video data from an open C3D file.
    
    This information is usually taken from the header record of the file.
    However, if the TRIAL:ACTUAL_START_FIELD and TRIAL:ACTUAL_END_FIELD parameters are present,
    the values from those parameters are used.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    int
        The last 3D frame number.

    """
    try:
        last_fr = itf.GetVideoFrame(1)
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None
    return np.int32(last_fr)

def get_num_frames(itf, log=False):
    """
    Get the total number of frames in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log: bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    int
        The total number of 3D frames.

    """
    last_fr = get_last_frame(itf, log)
    first_fr = get_first_frame(itf, log)
    if first_fr is None or last_fr is None:
        if log: logger.error('There is an error in getting either the first or the last frame number!')
        return None
    n_frs = last_fr-first_fr+1
    return np.int32(n_frs)

def check_frame_range_valid(itf, start_frame=None, end_frame=None, log=False):
    """
    Check the validity of input start and end frames.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    start_frame : int or None
        Input start frame.
    end_frame : int or None
        Input end frame.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int or None
        Valid start frame.
    int or None
        Valid end frame.

    """
    first_fr = get_first_frame(itf, log)
    last_fr = get_last_frame(itf, log)
    if first_fr is None or last_fr is None:
        if log: logger.error('There is an error in getting either the first or the last frame number!')
        return False, None, None
    if start_frame is None:
        start_fr = first_fr
    else:
        if start_frame < first_fr:
            if log: logger.error(f'"start_frame" should be equal or greater than {first_fr}!')
            return False, None, None
        start_fr = start_frame
    if end_frame is None:
        end_fr = last_fr
    else:
        if end_frame > last_fr:
            if log: logger.error(f'"end_frame" should be equal or less than {last_fr}!')
            return False, None, None
        end_fr = end_frame
    if not (start_fr < end_fr):
        if log: logger.error(f'Please provide a correct combination of "start_frame" and "end_frame"!')
        return False, None, None
    return True, start_fr, end_fr    

def get_video_fps(itf, log=False):
    """
    Return the 3D point sample rate in Hertz as read from the C3D file header.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    float
        Video frame rate in Hz from the header.

    """
    try:
        vid_fps = itf.GetVideoFrameRate()
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None    
    return np.float32(vid_fps)

def get_analog_video_ratio(itf, log=False):
    """
    Return the number of analog frames stored for each video frame in the C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    int
        The number of analog frames collected per video frame.

    """
    try:
        av_ratio = itf.GetAnalogVideoRatio()
    except pythoncom.com_error as err:
        if not (log and logger.isEnabledFor(logging.ERROR)):
            print(traceback.format_exc())
        if log: logger.error(err.excepinfo[2])
        return None    
    return np.int32(av_ratio)

def get_analog_fps(itf, log=False):
    """
    Return the analog sample rate in Hertz in the C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    float
        Analog sample rate in Hz.

    """
    vid_fps = get_video_fps(itf, log)
    av_ratio = get_analog_video_ratio(itf, log)
    if vid_fps is None or av_ratio is None:
        if log: logger.error('There is an error in getting necessary information!')
        return None    
    # return np.float32(get_video_fps(itf)*np.float32(get_analog_video_ratio(itf)))
    return np.float32(vid_fps*np.float32(av_ratio))

def get_video_frames(itf):
    """
    Return an integer-type numpy array that contains the video frame numbers between the start and the end frames.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.

    Returns
    -------
    frs : numpy array
        An integer-type numpy array of the video frame numbers.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)
    n_frs = end_fr-start_fr+1
    frs = np.linspace(start=start_fr, stop=end_fr, num=n_frs, dtype=np.int32)
    return frs

def get_analog_frames(itf):
    """
    Return a float-type numpy array that contains the analog frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.

    Returns
    -------
    frs : numpy array
        A float-type numpy array of the analog frame numbers.

    """
    av_ratio = get_analog_video_ratio(itf)
    start_fr = np.float32(get_first_frame(itf))
    end_fr = np.float32(get_last_frame(itf))+np.float32(av_ratio-1)/np.float32(av_ratio)
    analog_steps = get_num_frames(itf)*av_ratio
    frs = np.linspace(start=start_fr, stop=end_fr, num=analog_steps, dtype=np.float32)
    return frs

def get_video_times(itf, from_zero=True):
    """
    Return a float-type numpy array that contains the times corresponding to the video frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    from_zero : bool, optional
        Whether the return time array should start from zero or not. The default is True.

    Returns
    -------
    t : numpy array
        A float-type numpy array of the times corresponding to the video frame numbers.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)
    vid_fps = get_video_fps(itf)
    offset_fr = start_fr if from_zero else 0
    start_t = np.float32(start_fr-offset_fr)/vid_fps
    end_t = np.float32(end_fr-offset_fr)/vid_fps
    vid_steps = get_num_frames(itf)
    t = np.linspace(start=start_t, stop=end_t, num=vid_steps, dtype=np.float32)
    return t

def get_analog_times(itf, from_zero=True):
    """
    Return a float-type array that contains the times corresponding to the analog frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.

    Returns
    -------
    t : numpy array
        A float-type numpy array of the times corresponding to the analog frame numbers.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)
    vid_fps = get_video_fps(itf)
    analog_fps = get_analog_fps(itf)
    av_ratio = get_analog_video_ratio(itf)
    offset_fr = start_fr if from_zero else 0
    start_t = np.float32(start_fr-offset_fr)/vid_fps
    end_t = np.float32(end_fr-offset_fr)/vid_fps+np.float32(av_ratio-1)/analog_fps
    analog_steps = get_num_frames(itf)*av_ratio
    t = np.linspace(start=start_t, stop=end_t, num=analog_steps, dtype=np.float32)
    return t

def get_video_times_subset(itf, sel_masks):
    """
    Return a subset of the video frame time array.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sel_masks : list or numpy array
        list or numpy array of boolean for boolean array indexing.

    Returns
    -------
    numpy array
        A subset of the video frame time array.

    """
    return get_video_times(itf)[sel_masks]

def get_analog_times_subset(itf, sel_masks):
    """
    Return a subset of the analog frame time array.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sel_masks : list or numpy array
        list or numpy array of boolean for boolean array indexing.

    Returns
    -------
    numpy array
        A subset of the analog frame time array.

    """
    return get_analog_times(itf)[sel_masks]

def get_marker_names(itf, log=False):
    """
    Return a string-type list of the marker names from an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.
    
    Returns
    -------
    mkr_names : list or None
        A string-type list that contains the marker names.
        None if there is no POINT:LABELS parameter.
        None if there is no item in the POINT:LABELS parameter.
        
    """
    mkr_names = []
    idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
    if idx_pt_labels == -1:
        if log: logger.debug('No POINT:LABELS parameter!')
        return None
    n_pt_labels = itf.GetParameterLength(idx_pt_labels)
    if n_pt_labels < 1:
        if log: logger.debug('No item under POINT:LABELS parameter!')
        return None
    idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
    if idx_pt_used == -1:
        if log: logger.debug('No POINT:USED parameter!')
        return None
    n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
    if n_pt_used < 1:
        if log: logger.debug('POINT:USED value seems to be zero!')
        return None
    for i in range(n_pt_labels):
        if i < n_pt_used:
            mkr_names.append(itf.GetParameterValue(idx_pt_labels, i))
    return mkr_names

def get_marker_index(itf, mkr_name, log=False):
    """
    Return the index of given marker name in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    mkr_idx : int or None
        Marker index in the C3D file.
        None if there is no POINT:LABELS parameter.
        None if there is no item in the POINT:LABELS parameter.
        -1 if there is no corresponding marker with 'mkr_name' in the POINT:LABELS parameter.

    """
    idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
    if idx_pt_labels == -1:
        if log: logger.debug('No POINT:LABELS parameter!')
        return None
    n_pt_labels = itf.GetParameterLength(idx_pt_labels)
    if n_pt_labels < 1:
        if log: logger.debug('No item under POINT:LABELS parameter!')
        return None
    idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
    if idx_pt_used == -1:
        if log: logger.debug('No POINT:USED parameter!')
        return None
    n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
    if n_pt_used < 1:
        if log: logger.debug('POINT:USED value seems to be zero!')
        return None
    mkr_idx = -1
    for i in range(n_pt_labels):
        if i < n_pt_used:
            tgt_name = itf.GetParameterValue(idx_pt_labels, i)
            if tgt_name == mkr_name:
                mkr_idx = i
                break  
    if mkr_idx == -1:
        if log: logger.debug(f'No "{mkr_name}" marker exists!')
    return mkr_idx

def get_marker_unit(itf, log=False):
    """
    Return the unit of the marker coordinate values in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    unit : str or None
        The unit of the marker coordinate values.
        None if there is no POINT:UNITS parameter.
        None if there is no item in the POINT:UNITS parameter.

    """
    idx_pt_units = itf.GetParameterIndex('POINT', 'UNITS')
    if idx_pt_units == -1: 
        if log: logger.debug('No POINT:UNITS parameter!')
        return None
    n_items = itf.GetParameterLength(idx_pt_units)
    if n_items < 1: 
        if log: logger.debug('No item under POINT:UNITS parameter!')
        return None
    unit = itf.GetParameterValue(idx_pt_units, n_items-1)
    return unit

def get_marker_scale(itf, log=False):
    """
    Return the marker scale in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.         

    Returns
    -------
    scale : float or None
        The scale factor for marker coordinate values.
        None if there is no POINT:SCALE parameter.
        None if there is no item in the POINT:SCALE parameter.
    
    """
    idx_pt_scale = itf.GetParameterIndex('POINT', 'SCALE')
    if idx_pt_scale == -1:
        if log: logger.debug('No POINT:SCALE parameter!')
        return None
    n_items = itf.GetParameterLength(idx_pt_scale)
    if n_items < 1:
        if log: logger.debug('No item under POINT:SCALE parameter!')
        return None
    scale = np.float32(itf.GetParameterValue(idx_pt_scale, n_items-1))
    return scale

def get_marker_data(itf, mkr_name, blocked_nan=False, start_frame=None, end_frame=None, log=False):
    """
    Return the scaled marker coordinate values and the residuals in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    blocked_nan : bool, optional
        Whether to set the coordinates of blocked frames as nan. The default is False.
    start_frame: None or int, optional
        User-defined start frame.
    end_frame: None or int, optional
        User-defined end frame.
    log : bool, optional
        Whether to write logs or not. The default is False.         

    Returns
    -------
    mkr_data : numpy array or None
        2D numpy array (n, 4), where n is the number of frames in the output.
        For each row, the first three columns contains the x, y, z coordinates of the marker at each frame.
        For each row, The last (fourth) column contains the residual value.
        None if there is no corresponding marker name in the C3D file.
        
    """
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    n_frs = end_fr-start_fr+1
    mkr_data = np.full((n_frs, 4), np.nan, dtype=np.float32)
    for i in range(3):
        mkr_data[:,i] = np.array(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '1'), dtype=np.float32)
    mkr_data[:,3] = np.array(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
    if blocked_nan:
        mkr_null_masks = np.where(np.isclose(mkr_data[:,3], -1), True, False)
        mkr_data[mkr_null_masks,0:3] = np.nan 
    return mkr_data

def get_marker_pos(itf, mkr_name, blocked_nan=False, scaled=True, start_frame=None, end_frame=None, log=False):
    """
    Return a specific marker's coordinate values in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    blocked_nan : bool, optional
        Whether to set the coordinates of blocked frames as nan. The default is False.
    scaled : bool, optional
        Whether to return the scaled coordinate values or not. The default is True.        
    start_frame: None or int, optional
        User-defined start frame.
    end_frame: None or int, optional
        User-defined end frame.
    log : bool, optional
        Whether to write logs or not. The default is False.           

    Returns
    -------
    mkr_data : numpy array or None
        2D numpy array (n, 3), where n is the number of frames in the output.
        If 'blocked_nan' is set as True, then the corresponding row in the 'mkr_data' will be filled with nan.
        None if there is no corresponding marker name in the C3D file.
        
    Notes
    -----
    This is a wrapper function of GetPointDataEx() in the C3DServer SDK with 'byScaled' parameter as 1.
    
    """
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    n_frs = end_fr-start_fr+1
    mkr_scale = get_marker_scale(itf)
    is_c3d_float = mkr_scale < 0
    is_c3d_float2 = [False, True][itf.GetDataType()-1]
    if is_c3d_float != is_c3d_float2:
        if log: logger.debug(f'C3D data type is determined by POINT:SCALE parameter.')
    mkr_dtype = [[[np.int16, np.float32][is_c3d_float], np.float32][scaled], np.float32][blocked_nan]
    mkr_data = np.zeros((n_frs, 3), dtype=mkr_dtype)
    b_scaled = ['0', '1'][scaled]
    for i in range(3):
        mkr_data[:,i] = np.array(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, b_scaled), dtype=mkr_dtype)
    if blocked_nan:
        mkr_resid = np.array(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
        mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
        mkr_data[mkr_null_masks,:] = np.nan  
    return mkr_data

def get_marker_pos2(itf, mkr_name, blocked_nan=False, scaled=True, start_frame=None, end_frame=None, log=False):
    """
    Return a specific marker's coordinate values in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    blocked_nan : bool, optional
        Whether to set the coordinates of blocked frames as nan. The default is False.        
    scaled : bool, optional
        Whether to return the scaled coordinate values or not. The default is True.
    start_frame: None or int, optional
        User-defined start frame.
    end_frame: None or int, optional
        User-defined end frame.
    log : bool, optional
        Whether to write logs or not. The default is False.         

    Returns
    -------
    mkr_data : numpy array or None
        2D numpy array (n, 3), where n is the number of frames in the output.
        If 'blocked_nan' is set as True, then the corresponding row in the 'mkr_data' will be filled with nan.
        None if there is no corresponding marker name in the C3D file.
        
    Notes
    -----
    This is a wrapper function of GetPointDataEx() in the C3DServer SDK with 'byScaled' parameter as 0.        
    With this 'byScaled' as 0, GetPointDataEx() function will return un-scaled data if data is stored as integer format.
    Integer-format C3D files can be indentified by checking whether the scale value is positive or not. 
    For these integer-format C3D files, POINT:SCALE parameter will be used for scaling the coordinate values.
    This function returns the manual multiplication between un-scaled values from GetPointDataEx() and POINT:SCALE parameter.
    Ideally, get_marker_pos2() should return as same results as get_marker_pos() function.
    """
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    n_frs = end_fr-start_fr+1
    mkr_scale = get_marker_scale(itf)
    is_c3d_float = mkr_scale < 0
    is_c3d_float2 = [False, True][itf.GetDataType()-1]
    if is_c3d_float != is_c3d_float2:
        if log: logger.debug(f'C3D data type is determined by the POINT:SCALE parameter.')
    mkr_dtype = [[[np.int16, np.float32][is_c3d_float], np.float32][scaled], np.float32][blocked_nan]
    mkr_data = np.zeros((n_frs, 3), dtype=mkr_dtype)
    scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
    for i in range(3):
        if scaled:
            mkr_data[:,i] = np.array(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '0'), dtype=mkr_dtype)*scale_size
        else:
            mkr_data[:,i] = np.array(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '0'), dtype=mkr_dtype)
    if blocked_nan:    
        mkr_resid = np.array(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
        mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
        mkr_data[mkr_null_masks,:] = np.nan            
    return mkr_data

def get_marker_resid(itf, mkr_name, start_frame=None, end_frame=None, log=False):
    """
    Return the 3D residual values of a specified marker in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    start_frame: None or int, optional
        User-defined start frame.
    end_frame: None or int, optional
        User-defined end frame.
    log : bool, optional
        Whether to write logs or not. The default is False.           

    Returns
    -------
    mkr_resid : numpy array or None
        1D numpy array (n,), where n is the number of frames in the output.

    """
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    mkr_resid = np.array(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
    return mkr_resid

def get_analog_names(itf, log=False):
    """
    Return a string list of the analog channel names in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.
        
    Returns
    -------
    sig_names : list
        String list that contains the analog channel names.

    """
    sig_names = []
    idx_anl_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    if idx_anl_labels == -1:
        if log: logger.debug('No ANALOG:LABELS parameter!')
        return None
    n_anl_labels = itf.GetParameterLength(idx_anl_labels)
    if n_anl_labels < 1:
        if log: logger.debug('No item under ANALOG:LABELS parameter!')
        return None
    idx_anl_used = itf.GetParameterIndex('ANALOG', 'USED')
    if idx_anl_used == -1:
        if log: logger.debug('No ANALOG:USED parameter!')
        return None        
    n_anl_used = itf.GetParameterValue(idx_anl_used, 0)    
    for i in range(n_anl_labels):
        if i < n_anl_used:
            sig_names.append(itf.GetParameterValue(idx_anl_labels, i))
    return sig_names

def get_analog_index(itf, sig_name, log=False):
    """
    Get the index of analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig_idx : int
        Index of the analog channel.

    """
    idx_anl_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    if idx_anl_labels == -1:
        if log: logger.debug('No ANALOG:LABELS parameter!')
        return None
    n_anl_labels = itf.GetParameterLength(idx_anl_labels)
    if n_anl_labels < 1:
        if log: logger.debug('No item under ANALOG:LABELS parameter!')
        return None
    idx_anl_used = itf.GetParameterIndex('ANALOG', 'USED')
    if idx_anl_used == -1:
        if log: logger.debug('No ANALOG:USED parameter!')
        return None
    n_anl_used = itf.GetParameterValue(idx_anl_used, 0)    
    sig_idx = -1    
    for i in range(n_anl_labels):
        if i < n_anl_used:
            tgt_name = itf.GetParameterValue(idx_anl_labels, i)
            if tgt_name == sig_name:
                sig_idx = i
                break        
    if sig_idx == -1:
        if log: logger.debug(f'No "{sig_name}" analog channel in the open file!')
    return sig_idx

def get_analog_gen_scale(itf, log=False):
    """
    Return the general (common) scaling factor for analog channels in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    gen_scale : float or None
        The general (common) scaling factor for analog channels.
        None if there is no ANALOG:GEN_SCALE parameter in the C3D file.
        None if there is no item in the ANALOG:GEN_SCALE parameter.
        
    """
    par_idx = itf.GetParameterIndex('ANALOG', 'GEN_SCALE')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:GEN_SCALE parameter!')
        return None
    n_items = itf.GetParameterLength(par_idx)
    if n_items < 1:
        if log: logger.debug('No item under ANALOG:GEN_SCALE parameter!')
        return None
    gen_scale = np.float32(itf.GetParameterValue(par_idx, n_items-1))
    return gen_scale

def get_analog_format(itf, log=False):
    """
    Return the format of analog channels in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    sig_format : str or None
        Format of the analog channels.
        None if there is no ANALOG:FORMAT parameter in the C3D file.
        None if there is no item in the ANALOG:FORMAT parameter.
        
    """
    par_idx = itf.GetParameterIndex('ANALOG', 'FORMAT')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:FORMAT parameter!')
        return None
    n_items = itf.GetParameterLength(par_idx)
    if n_items < 1:
        if log: logger.debug('No item under ANALOG:FORMAT parameter!')
        return None    
    sig_format = itf.GetParameterValue(par_idx, n_items-1)
    return sig_format

def get_analog_unit(itf, sig_name, log=False):
    """
    Return the unit of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig_unit : str or None
        Analog channel unit.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    par_idx = itf.GetParameterIndex('ANALOG', 'UNITS')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:UNITS parameter!')
        return None
    sig_unit = itf.GetParameterValue(par_idx, sig_idx)
    return sig_unit
    
def get_analog_scale(itf, sig_name, log=False):
    """
    Return the scale of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig_scale : float or None
        Analog channel scale.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    par_idx = itf.GetParameterIndex('ANALOG', 'SCALE')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:SCALE parameter!')
        return None
    sig_scale = np.float32(itf.GetParameterValue(par_idx, sig_idx))
    return sig_scale
    
def get_analog_offset(itf, sig_name, log=False):
    """
    Return the offset of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig_offset : int or None
        Analog channel offset.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    par_idx = itf.GetParameterIndex('ANALOG', 'OFFSET')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:OFFSET parameter!')
        return None
    sig_format = get_analog_format(itf)
    is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
    par_dtype = [np.int16, np.uint16][is_sig_unsigned]
    sig_offset = par_dtype(itf.GetParameterValue(par_idx, sig_idx))
    return sig_offset
            
def get_analog_data_unscaled(itf, sig_name, start_frame=None, end_frame=None, log=False):
    """
    Return the unscaled value of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    start_frame : int or None, optional
        Start frame number. The default is None.
    end_frame : int or None, optional
        End frame number. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig : numpy array or None
        Analog channel value.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    sig_format = get_analog_format(itf)
    is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')        
    mkr_scale = get_marker_scale(itf)
    is_c3d_float = mkr_scale < 0
    is_c3d_float2 = [False, True][itf.GetDataType()-1]
    if is_c3d_float != is_c3d_float2:
        if log: logger.debug(f'C3D data type is determined by the POINT:SCALE parameter.')
    sig_dtype = [[np.int16, np.uint16][is_sig_unsigned], np.float32][is_c3d_float]
    sig = np.array(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=sig_dtype)
    return sig

def get_analog_data_scaled(itf, sig_name, start_frame=None, end_frame=None, log=False):
    """
    Return the scale value of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    start_frame : int or None, optional
        Start frame number. The default is None.
    end_frame : int or None, optional
        End frame number. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig : numpy array or None
        Analog channel value.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    sig = np.array(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '1', 0, 0, '0'), dtype=np.float32)
    return sig

def get_analog_data_scaled2(itf, sig_name, start_frame=None, end_frame=None, log=False):
    """
    Return the scale value of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    start_frame : int or None, optional
        Start frame number. The default is None.
    end_frame : int or None, optional
        End frame number. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    sig : numpy array or None
        Analog channel value.

    """
    sig_idx = get_analog_index(itf, sig_name, log)
    if sig_idx == -1 or sig_idx is None: return None
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log)
    if not fr_check: return None
    gen_scale = get_analog_gen_scale(itf)
    sig_scale = get_analog_scale(itf, sig_name)
    sig_offset = np.float32(get_analog_offset(itf, sig_name))
    sig = (np.array(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-sig_offset)*sig_scale*gen_scale
    return sig

def get_dict_header(itf):
    """
    Return the summarization of the C3D header information.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.

    Returns
    -------
    dict_header : dict
        Dictionary of the C3D header information.

    """
    dict_file_type = {1:'INTEL', 2:'DEC', 3:'SGI'}
    dict_data_type = {1:'INTEGER', 2:'REAL'}
    dict_header = {}
    dict_header['FILE_TYPE'] = dict_file_type.get(itf.GetFileType(), None)
    dict_header['DATA_TYPE'] = dict_data_type.get(itf.GetDataType(), None)
    dict_header['NUM_3D_POINTS'] = np.int32(itf.GetNumber3DPoints())
    dict_header['NUM_ANALOG_CHANNELS'] = np.int32(itf.GetAnalogChannels())
    dict_header['FIRST_FRAME'] = np.int32(itf.GetVideoFrameHeader(0))
    dict_header['LAST_FRAME'] = np.int32(itf.GetVideoFrameHeader(1))
    dict_header['START_RECORD'] = np.int32(itf.GetStartingRecord())
    dict_header['VIDEO_FRAME_RATE'] = np.float32(itf.GetVideoFrameRate())
    dict_header['ANALOG_VIDEO_RATIO'] = np.int32(itf.GetAnalogVideoRatio())
    dict_header['ANALOG_FRAME_RATE'] = np.float32(itf.GetVideoFrameRate()*itf.GetAnalogVideoRatio())
    dict_header['MAX_INTERPOLATION_GAP'] = np.int32(itf.GetMaxInterpolationGap())
    dict_header['3D_SCALE_FACTOR'] = np.float32(itf.GetHeaderScaleFactor())
    return dict_header

def get_dict_groups(itf, tgt_grp_names=None):
    """
    Return the dictionary of the groups.

    All the values in the dictionary structure are numpy arrays except the values of single scalar.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_grp_names: str or tuple, optional
        Target group names to extract. The default is None.
    
    Returns
    -------
    dict_grps : dict
        Dictionary of the C3D header information.

    """
    dict_dtype = {-1:str, 1:np.int8, 2:np.int32, 4:np.float32}
    dict_grps = {}
    dict_grp_names = {}
    n_grps = itf.GetNumberGroups()
    for i in range(n_grps):
        grp_name = itf.GetGroupName(i)
        if (tgt_grp_names is not None) and (grp_name not in tgt_grp_names): continue
        grp_number = itf.GetGroupNumber(i)
        dict_grp_names.update({np.absolute(grp_number, dtype=np.int): grp_name})
        dict_grps[grp_name] = {}
    n_params = itf.GetNumberParameters()
    for i in range(n_params):
        par_num = itf.GetParameterNumber(i)
        grp_name = dict_grp_names.get(par_num, None)
        if grp_name is None: continue
        if (tgt_grp_names is not None) and (grp_name not in tgt_grp_names): continue
        par_name = itf.GetParameterName(i)
        par_len = itf.GetParameterLength(i)
        par_type = itf.GetParameterType(i)
        data_type = dict_dtype.get(par_type, None)
        par_data = []
        if grp_name=='ANALOG' and par_name=='OFFSET':
            sig_format = get_analog_format(itf)
            is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
            pre_dtype = [np.int16, np.uint16][is_sig_unsigned]
            for j in range(par_len):
                par_data.append(pre_dtype(itf.GetParameterValue(i, j)))
        else:
            for j in range(par_len):
                par_data.append(itf.GetParameterValue(i, j))
        dict_grps[grp_name][par_name] = data_type(par_data[0]) if len(par_data)==1 else np.asarray(par_data, dtype=data_type)
    return dict_grps

def get_dict_markers(itf, blocked_nan=False, resid=False, mask=False, desc=False, frame=False, time=False, tgt_mkr_names=None, log=False):
    """
    Get the dictionary of marker information.
    
    All marker position values will be scaled.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    blocked_nan : bool, optional
        Whether to set the coordinates of blocked frames as nan. The default is False.
    resid : bool, optional
        Whether to include the residual values of markers. The default is False.
    mask : bool, optional
        Whether to include the mask information of markers. The default is False.
    desc : bool, optional
        Whether to include the descriptions of markers. The default is False.        
    frame : bool, optional
        Whether to include the frame array. The default is False.        
    time : bool, optional
        Whether to include the time array. The default is False.
    tgt_mkr_names : list or tuple, optional
        Specific target marker names to extract. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict_pts : dictionary
        Dictionary of marker information.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)    
    n_frs = end_fr-start_fr+1
    idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
    if idx_pt_labels == -1: idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS1')
    if idx_pt_labels == -1: idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS2')
    if idx_pt_labels == -1: idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS3')
    if idx_pt_labels == -1:
        if log: logger.debug('No POINT:LABELS parameter!')
        return None
    n_pt_labels = itf.GetParameterLength(idx_pt_labels)
    if n_pt_labels < 1:
        if log: logger.debug('No item under POINT:LABELS parameter!')
        return None
    idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
    if idx_pt_used == -1:
        if log: logger.debug('No POINT:USED parameter!')
        return None        
    n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
    if n_pt_used < 1:
        if log: logger.debug('POINT:USED is zero!')
        return None
    idx_pt_desc = itf.GetParameterIndex('POINT', 'DESCRIPTIONS')
    if idx_pt_desc == -1:
        if log: logger.debug('No POINT:DESCRIPTIONS parameter!')
        n_pt_desc = 0
    else:
        n_pt_desc = itf.GetParameterLength(idx_pt_desc)
    dict_pts = {}
    mkr_names = []
    mkr_descs = []
    dict_pts.update({'DATA':{}})
    dict_pts['DATA'].update({'POS':{}})
    if resid: dict_pts['DATA'].update({'RESID': {}})
    if mask: dict_pts['DATA'].update({'MASK': {}})
    for i in range(n_pt_labels):
        if i < n_pt_used:
            mkr_name = itf.GetParameterValue(idx_pt_labels, i)
            if (tgt_mkr_names is not None) and (mkr_name not in tgt_mkr_names): continue
            mkr_names.append(mkr_name)
            mkr_data = np.zeros((n_frs, 3), dtype=np.float32)
            for j in range(3):
                mkr_data[:,j] = np.array(itf.GetPointDataEx(i, j, start_fr, end_fr, '1'), dtype=np.float32)
            if blocked_nan or resid:
                mkr_resid = np.array(itf.GetPointResidualEx(i, start_fr, end_fr), dtype=np.float32)
            if blocked_nan:
                mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
                mkr_data[mkr_null_masks,:] = np.nan
            dict_pts['DATA']['POS'].update({mkr_name: mkr_data})
            if resid:
                dict_pts['DATA']['RESID'].update({mkr_name: mkr_resid})
            if mask:
                mkr_mask = np.array(itf.GetPointMaskEx(i, start_fr, end_fr), dtype=str)
                dict_pts['DATA']['MASK'].update({mkr_name: mkr_mask})
            if desc:
                if i < n_pt_desc:
                    mkr_descs.append(itf.GetParameterValue(idx_pt_desc, i))
                else:
                    mkr_descs.append('')
    dict_pts.update({'LABELS': np.array(mkr_names, dtype=str)})
    idx_pt_rate = itf.GetParameterIndex('POINT', 'RATE')
    if idx_pt_rate != -1:
        n_pt_rate = itf.GetParameterLength(idx_pt_rate)
        if n_pt_rate == 1:
            rate = np.float32(itf.GetParameterValue(idx_pt_rate, 0))
            dict_pts.update({'RATE': rate})
    idx_pt_units = itf.GetParameterIndex('POINT', 'UNITS')
    if idx_pt_units != -1:
        n_pt_units = itf.GetParameterLength(idx_pt_units)
        if n_pt_units == 1:
            unit = itf.GetParameterValue(idx_pt_units, 0)
            dict_pts.update({'UNITS': unit})
    if desc:
        if idx_pt_desc != -1:
            dict_pts.update({'DESCRIPTIONS': np.array(mkr_descs, dtype=str)})
    if frame: dict_pts.update({'FRAME': get_video_frames(itf)})
    if time: dict_pts.update({'TIME': get_video_times(itf)})
    return dict_pts

def get_dict_forces(itf, desc=False, frame=False, time=False, log=False):
    """
    Get the dictionary of forces.
    
    All force (analog) values will be scaled.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    desc : bool, optional
        Whether to include the descriptions of forces. The default is False.        
    frame : bool, optional
        Whether to include the frame array. The default is False.          
    time : bool, optional
        Whether to include the time array. The default is False.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict_forces : dictionary
        Dictionary of force information.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)
    idx_force_used = itf.GetParameterIndex('FORCE_PLATFORM', 'USED')
    if idx_force_used == -1: 
        if log: logger.debug(f'FORCE_PLATFORM:USED parameter does not exist!')
        return None
    n_force_used = itf.GetParameterValue(idx_force_used, 0)
    if n_force_used < 1:
        if log: logger.debug(f'FORCE_PLATFORM:USED is zero!')
        return None
    idx_force_chs = itf.GetParameterIndex('FORCE_PLATFORM', 'CHANNEL')
    if idx_force_chs == -1: 
        if log: logger.debug(f'FORCE_PLATFORM:CHANNEL parameter does not exist!')
        return None
    idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    if idx_analog_labels == -1:
        if log: logger.debug('No ANALOG:LABELS parameter!')
        return None       
    idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
    if idx_analog_scale == -1:
        if log: logger.debug('No ANALOG:SCALE parameter!')
        return None     
    idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
    if idx_analog_offset == -1:
        if log: logger.debug('No ANALOG:OFFSET parameter!')
        return None
    idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
    if idx_analog_units == -1:
        if log: logger.debug('No ANALOG:UNITS parameter!')
        n_analog_units = 0
    else:
        n_analog_units = itf.GetParameterLength(idx_analog_units)
    idx_analog_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
    if idx_analog_desc == -1:
        if log: logger.debug('No ANALOG:DESCRIPTIONS parameter!')
        n_analog_desc = 0
    else:
        n_analog_desc = itf.GetParameterLength(idx_analog_desc)
    gen_scale = get_analog_gen_scale(itf)
    sig_format = get_analog_format(itf)
    is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
    offset_dtype = [np.int16, np.uint16][is_sig_unsigned]
    dict_forces = {}
    force_names = []
    force_units = []
    force_descs = []
    dict_forces.update({'DATA':{}})
    n_force_chs = itf.GetParameterLength(idx_force_chs)
    for i in range(n_force_chs):
        ch_idx = itf.GetParameterValue(idx_force_chs, i)-1
        ch_name = itf.GetParameterValue(idx_analog_labels, ch_idx)
        force_names.append(ch_name)
        ch_scale = np.float32(itf.GetParameterValue(idx_analog_scale, ch_idx))
        ch_offset = np.float32(offset_dtype(itf.GetParameterValue(idx_analog_offset, ch_idx)))
        ch_val = (np.array(itf.GetAnalogDataEx(ch_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-ch_offset)*ch_scale*gen_scale
        dict_forces['DATA'].update({ch_name: ch_val})
        if ch_idx < n_analog_units:
            force_units.append(itf.GetParameterValue(idx_analog_units, ch_idx))
        else:
            force_units.append('')
        if desc:
            if ch_idx < n_analog_desc:
                force_descs.append(itf.GetParameterValue(idx_analog_desc, ch_idx))
            else:
                force_descs.append('')
    dict_forces.update({'LABELS': np.array(force_names, dtype=str)})
    idx_analog_rate = itf.GetParameterIndex('ANALOG', 'RATE')
    if idx_analog_rate != -1:
        n_analog_rate = itf.GetParameterLength(idx_analog_rate)
        if n_analog_rate == 1:
            dict_forces.update({'RATE': np.float32(itf.GetParameterValue(idx_analog_rate, 0))})
    if idx_analog_units != -1:
        dict_forces.update({'UNITS': np.array(force_units, dtype=str)})
    if desc:
        if idx_analog_desc != -1:
            dict_forces.update({'DESCRIPTIONS': np.array(force_descs, dtype=str)})
    if frame: dict_forces.update({'FRAME': get_analog_frames(itf)})
    if time: dict_forces.update({'TIME': get_analog_times(itf)})
    return dict_forces

def get_dict_analogs(itf, desc=False, frame=False, time=False, excl_forces=True, log=False):
    """
    Get the dictionary of analogs.
    
    All analog channel values will be scaled.    

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    desc : bool, optional
        Whether to include the descriptions of analogs. The default is False.        
    frame : bool, optional
        Whether to include the frame array. The default is False.          
    time : bool, optional
        Whether to include the time array. The default is False.
    excl_forces : bool, optional
        Whether to exclude forces in the output or not. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict_analogs : dictionary
        Dictionary of analog information.

    """
    start_fr = get_first_frame(itf)
    end_fr = get_last_frame(itf)
    n_force_chs = 0
    idx_force_chs = itf.GetParameterIndex('FORCE_PLATFORM', 'CHANNEL')
    if idx_force_chs == -1: 
        if log: logger.debug(f'FORCE_PLATFORM:CHANNEL parameter does not exist!')
        n_force_chs = 0
    else:
        n_force_chs = itf.GetParameterLength(idx_force_chs)
    force_ch_idx = []
    if excl_forces:
        for i in range(n_force_chs):
            ch_idx = itf.GetParameterValue(idx_force_chs, i)-1
            force_ch_idx.append(ch_idx)
    idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    if idx_analog_labels == -1:
        if log: logger.debug('No ANALOG:LABELS parameter!')
        return None
    n_analog_labels = itf.GetParameterLength(idx_analog_labels)
    if n_analog_labels < 1:
        if log: logger.debug('No item under ANALOG:LABELS parameter!')
        return None    
    idx_analog_used = itf.GetParameterIndex('ANALOG', 'USED')
    if idx_analog_used == -1:
        if log: logger.debug('No ANALOG:USED parameter!')
        return None
    n_analog_used = itf.GetParameterValue(idx_analog_used, 0)
    if n_analog_used < 1:
        if log: logger.debug(f'ANALOG:USED is zero!')
        return None    
    idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
    if idx_analog_scale == -1:
        if log: logger.debug('No ANALOG:SCALE parameter!')
        return None       
    idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
    if idx_analog_offset == -1:
        if log: logger.debug('No ANALOG:OFFSET parameter!')
        return None
    idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
    if idx_analog_units == -1:
        if log: logger.debug('No ANALOG:UNITS parameter!')
        n_analog_units = 0
    else:
        n_analog_units = itf.GetParameterLength(idx_analog_units)
    idx_analog_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
    if idx_analog_desc == -1:
        if log: logger.debug('No ANALOG:DESCRIPTIONS parameter!')
        n_analog_desc = 0
    else:
        n_analog_desc = itf.GetParameterLength(idx_analog_desc)
    gen_scale = get_analog_gen_scale(itf)
    sig_format = get_analog_format(itf)
    is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
    offset_dtype = [np.int16, np.uint16][is_sig_unsigned]    
    dict_analogs = {}
    analog_names = []
    analog_units = []
    analog_descs = []
    dict_analogs.update({'DATA':{}})
    for i in range(n_analog_labels):
        if i < n_analog_used:
            if i in force_ch_idx: continue
            sig_name = itf.GetParameterValue(idx_analog_labels, i)
            analog_names.append(sig_name)
            sig_scale = np.float32(itf.GetParameterValue(idx_analog_scale, i))
            sig_offset = np.float32(offset_dtype(itf.GetParameterValue(idx_analog_offset, i)))
            sig_val = (np.array(itf.GetAnalogDataEx(i, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-sig_offset)*sig_scale*gen_scale
            dict_analogs['DATA'].update({sig_name: sig_val})
            if i < n_analog_units:
                analog_units.append(itf.GetParameterValue(idx_analog_units, i))
            else:
                analog_units.append('')
            if desc:
                if i < n_analog_desc:
                    analog_descs.append(itf.GetParameterValue(idx_analog_desc, i))
                else:
                    analog_descs.append('')
    dict_analogs.update({'LABELS': np.array(analog_names, dtype=str)})
    idx_analog_rate = itf.GetParameterIndex('ANALOG', 'RATE')
    if idx_analog_rate != -1:
        n_analog_rate = itf.GetParameterLength(idx_analog_rate)
        if n_analog_rate == 1:
            dict_analogs.update({'RATE': np.float32(itf.GetParameterValue(idx_analog_rate, 0))})
    if idx_analog_units != -1:
        dict_analogs.update({'UNITS': np.array(analog_units, dtype=str)})
    if desc:
        if idx_analog_desc != -1:
            dict_analogs.update({'DESCRIPTIONS': np.array(analog_descs, dtype=str)})
    if frame: dict_analogs.update({'FRAME': get_analog_frames(itf)})
    if time: dict_analogs.update({'TIME': get_analog_times(itf)})
    return dict_analogs
    
def change_marker_name(itf, mkr_name_old, mkr_name_new, log=False):
    """
    Change the name of a marker.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name_old : str
        Old marker name.
    mkr_name_new : str
        New marker name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    mkr_idx = get_marker_index(itf, mkr_name_old, log)
    if mkr_idx == -1 or mkr_idx is None: return False
    par_idx = itf.GetParameterIndex('POINT', 'LABELS')
    if par_idx == -1:
        if log: logger.debug('No POINT:LABELS parameter!')
        return False
    ret = itf.SetParameterValue(par_idx, mkr_idx, mkr_name_new)
    if log:
        logger.info(f'Changing of the marker name from "{mkr_name_old}" to "{mkr_name_new}" is {["not performed.", "performed."][ret]}')
    return [False, True][ret]

def change_analog_name(itf, sig_name_old, sig_name_new, log=False):
    """
    Change the name of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name_old : str
        Old analog channel name.
    sig_name_new : str
        New analog channel name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    sig_idx = get_analog_index(itf, sig_name_old, log)
    if sig_idx == -1 or sig_idx is None: return False
    par_idx = itf.GetParameterIndex('ANALOG', 'LABELS')
    if par_idx == -1:
        if log: logger.debug('No ANALOG:LABELS parameter!')
        return False        
    ret = itf.SetParameterValue(par_idx, sig_idx, sig_name_new)
    if log:
        logger.info(f'Changing of the signal name from "{sig_name_old}" to "{sig_name_new}" is {["not performed.", "performed."][ret]}')    
    return [False, True][ret]

def add_marker(itf, mkr_name, mkr_coords, mkr_resid=None, mkr_desc=None, log=False):
    """
    Add a new marker into an open C3D file.
    
    This function only works normally if 'POINT:USED' is as same as the number of items under 'POINT:LABELS'.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        A new marker name.
    mkr_coords : numpy array
        A numpy array of new marker coordinates.
    mkr_resid : numpy array or None, optional
        A numpy array of new marker residuals. The default is None.
    mkr_desc: str or None, optional
        Description of a new marker.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True of False.

    """
    if log: logger.debug(f'Start adding a new "{mkr_name}" marker ...')
    start_fr = get_first_frame(itf)
    n_frs = get_num_frames(itf)
    if not (mkr_coords.ndim == 2 and mkr_coords.shape[0] == n_frs and mkr_coords.shape[1] == 3):
        if log: logger.error('The dimension of the input marker coordinates are not valid!')
        return False
    if mkr_resid is not None:
        if not (mkr_resid.ndim == 1 and mkr_resid.shape[0] == n_frs):
            if log: logger.error('The dimension of the input marker residuals are not valid!')
            return False
    ret = 0
    # Check the value 'POINT:USED'
    par_idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
    n_pt_used_before = itf.GetParameterValue(par_idx_pt_used, 0)
    # Check the value 'POINT:LABELS'
    par_idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
    n_pt_labels_before = itf.GetParameterLength(par_idx_pt_labels)
    # Skip if 'POINT:USED' and 'POINT:LABELS' have different numbers
    if n_pt_used_before != n_pt_labels_before:
        if log: logger.error('This function only works if POINT:USED is as same as the number of items under POINT:LABELS!')
        return False
    # Add an parameter to the 'POINT:LABELS' section
    # par_idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
    ret = itf.AddParameterData(par_idx_pt_labels, 1)
    n_pt_labels = itf.GetParameterLength(par_idx_pt_labels)
    variant = win32.VARIANT(pythoncom.VT_BSTR, np.string_(mkr_name))
    ret = itf.SetParameterValue(par_idx_pt_labels, n_pt_labels-1, variant)
    # Add a null parameter in the 'POINT:DESCRIPTIONS' section
    par_idx_pt_desc = itf.GetParameterIndex('POINT', 'DESCRIPTIONS')
    ret = itf.AddParameterData(par_idx_pt_desc, 1)
    n_pt_desc = itf.GetParameterLength(par_idx_pt_desc)
    mkr_desc_adjusted = '' if mkr_desc is None else mkr_desc
    variant = win32.VARIANT(pythoncom.VT_BSTR, np.string_(mkr_desc_adjusted))
    ret = itf.SetParameterValue(par_idx_pt_desc, n_pt_desc-1, variant)
    # Add a marker
    new_mkr_idx = itf.AddMarker()
    n_mkrs = itf.GetNumber3DPoints()
    mkr_null_mask = np.any(np.isnan(mkr_coords), axis=1)
    mkr_resid_adjusted = np.zeros((n_frs, ), dtype=np.float32) if mkr_resid is None else np.array(mkr_resid, dtype=np.float32)
    mkr_resid_adjusted[mkr_null_mask] = -1
    mkr_masks = np.array(['0000000']*n_frs, dtype = np.string_)
    mkr_scale = get_marker_scale(itf)
    is_c3d_float = mkr_scale < 0
    is_c3d_float2 = [False, True][itf.GetDataType()-1]
    if is_c3d_float != is_c3d_float2:
        if log: logger.debug('C3D data type is determined by the POINT:SCALE parameter.')
    mkr_dtype = [np.int16, np.float32][is_c3d_float]    
    scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
    if is_c3d_float:
        mkr_coords_unscaled = np.asarray(np.nan_to_num(mkr_coords), dtype=mkr_dtype)
    else:
        mkr_coords_unscaled = np.asarray(np.round(np.nan_to_num(mkr_coords)/scale_size), dtype=mkr_dtype)
    dtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_c3d_float]
    dtype_arr = pythoncom.VT_ARRAY|dtype
    for i in range(3):
        variant = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i])
        ret = itf.SetPointDataEx(n_mkrs-1, i, start_fr, variant)
    variant = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, mkr_resid_adjusted)
    ret = itf.SetPointDataEx(n_mkrs-1, 3, start_fr, variant)
    variant = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_BSTR, mkr_masks)
    ret = itf.SetPointDataEx(n_mkrs-1, 4, start_fr, variant)        
    var_const = win32.VARIANT(dtype, 1)
    for i in range(3):
        for idx, val in enumerate(mkr_coords_unscaled[:,i]):
            if val == 1:
                ret = itf.SetPointData(n_mkrs-1, i, start_fr+idx, var_const)
    return [False, True][ret]
    # Increase the value 'POINT:USED' by the 1
    par_idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
    n_pt_used_after = itf.GetParameterValue(par_idx_pt_used, 0)
    if n_pt_used_after != (n_pt_used_before+1):
        if log: log.debug('POINT:USED was not properly updated so that manual update will be executed.')
        ret = itf.SetParameterValue(par_idx_pt_used, 0, (n_pt_used_before+1))
    return [False, True][ret]

def add_analog(itf, sig_name, sig_value, sig_unit, sig_scale=1.0, sig_offset=0, sig_gain=0, sig_desc=None, log=False):
    """
    Add a new analog signal into an open C3D file.
    
    This function only works normally if 'ANALOG:USED' is as same as the number of items under 'ANALOG:LABELS'.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        A new analog channel name.
    sig_value : numpy array
        A new analog channel value.
    sig_unit : str
        A new analog channel unit.        
    sig_scale : float, optional
        A new analog channel scale. The default is 1.0.
    sig_offset : int, optional
        A new analog channel offset. The default is 0.
    sig_gain : int, optional
        A new analog channel gain. The default is 0.
    sig_desc : str, optional
        A new analog channel description. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    if log: logger.debug(f'Start adding a new "{sig_name}" analog channel ...')
    start_fr = get_first_frame(itf)
    n_frs = get_num_frames(itf)
    av_ratio = get_analog_video_ratio(itf)
    if sig_value.ndim!=1 or sig_value.shape[0]!=(n_frs*av_ratio):
        if log: logger.error('The dimension of the input is not compatible!')
        return False
    # Check 'ANALOG:USED'
    n_idx_analog_used = itf.GetParameterIndex('ANALOG', 'USED')
    n_cnt_analog_used_before = itf.GetParameterValue(n_idx_analog_used, 0) 
    # Check 'ANALOG:LABELS'
    n_idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    n_cnt_analog_labels_before = itf.GetParameterLength(n_idx_analog_labels)
    # Skip if 'ANALOG:USED' and 'ANALOG:LABELS' have different numbers
    if n_cnt_analog_used_before != n_cnt_analog_labels_before:
        if log: logger.error('This function only works if ANALOG:USED is as same as the number of items under ANALOG:LABELS!')
        return False    
    # Add an parameter to the 'ANALOG:LABELS' section
    n_idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
    ret = itf.AddParameterData(n_idx_analog_labels, 1)
    n_cnt_analog_labels = itf.GetParameterLength(n_idx_analog_labels)
    ret = itf.SetParameterValue(n_idx_analog_labels, n_cnt_analog_labels-1, win32.VARIANT(pythoncom.VT_BSTR, sig_name))
    # Add an parameter to the 'ANALOG:UNITS' section
    n_idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
    ret = itf.AddParameterData(n_idx_analog_units, 1)
    n_cnt_analog_units = itf.GetParameterLength(n_idx_analog_units)
    ret = itf.SetParameterValue(n_idx_analog_units, n_cnt_analog_units-1, win32.VARIANT(pythoncom.VT_BSTR, sig_unit))      
    # Add an parameter to the 'ANALOG:SCALE' section
    n_idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
    ret = itf.AddParameterData(n_idx_analog_scale, 1)
    n_cnt_analog_scale = itf.GetParameterLength(n_idx_analog_scale)
    ret = itf.SetParameterValue(n_idx_analog_scale, n_cnt_analog_scale-1, win32.VARIANT(pythoncom.VT_R4, sig_scale))
    # Add an parameter to the 'ANALOG:OFFSET' section
    n_idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
    ret = itf.AddParameterData(n_idx_analog_offset, 1)
    n_cnt_analog_offset = itf.GetParameterLength(n_idx_analog_offset)
    sig_format = get_analog_format(itf)
    is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
    sig_offset_comtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_sig_unsigned]
    sig_offset_dtype = [np.int16, np.uint16][is_sig_unsigned]
    ret = itf.SetParameterValue(n_idx_analog_offset, n_cnt_analog_offset-1, win32.VARIANT(sig_offset_comtype, sig_offset))
    # Check for 'ANALOG:GAIN' section and add 0 if it exists
    n_idx_analog_gain = itf.GetParameterIndex('ANALOG', 'GAIN')
    if n_idx_analog_gain != -1:
        ret = itf.AddParameterData(n_idx_analog_gain, 1)
        n_cnt_analog_gain = itf.GetParameterLength(n_idx_analog_gain)
        ret = itf.SetParameterValue(n_idx_analog_gain, n_cnt_analog_gain-1, win32.VARIANT(pythoncom.VT_I2, sig_gain))    
    # Add an parameter to the 'ANALOG:DESCRIPTIONS' section
    n_idx_analog_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
    ret = itf.AddParameterData(n_idx_analog_desc, 1)
    n_cnt_analog_desc = itf.GetParameterLength(n_idx_analog_desc)
    sig_desc_in = sig_name if sig_desc is None else sig_desc
    ret = itf.SetParameterValue(n_idx_analog_desc, n_cnt_analog_desc-1, win32.VARIANT(pythoncom.VT_BSTR, sig_desc_in))
    # Create an analog channel
    n_idx_new_analog_ch = itf.AddAnalogChannel()
    n_cnt_analog_chs = itf.GetAnalogChannels()
    gen_scale = get_analog_gen_scale(itf)
    sig_value_unscaled = np.asarray(sig_value, dtype=np.float32)/(np.float32(sig_scale)*gen_scale)+np.float32(sig_offset_dtype(sig_offset))
    ret = itf.SetAnalogDataEx(n_idx_new_analog_ch, start_fr, win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, sig_value_unscaled))
    # Increase the value 'ANALOG:USED' by the 1
    n_idx_analog_used = itf.GetParameterIndex('ANALOG', 'USED')
    n_cnt_analog_used_after = itf.GetParameterValue(n_idx_analog_used, 0)
    if n_cnt_analog_used_after != (n_cnt_analog_used_before+1):
        if log: log.debug('ANALOG:USED was not properly updated so that manual update will be executed.')
        ret = itf.SetParameterValue(n_idx_analog_used, 0, (n_cnt_analog_used_before+1))
    return [False, True][ret]

def delete_frames(itf, start_frame, num_frames, log=False):
    """
    Delete specified frames in an open C3D file.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    start_frame : int
        Start frame number.
    num_frames : int
        Number of frames to be deleted.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    n_frs_updated : int
        Number of the remaining frames in the C3D file.

    """
    if start_frame < get_first_frame(itf):
        if log: logger.error(f'Given start frame number should be equal or greater than {get_first_frame(itf)} for the open file!')
        return None
    elif start_frame >= get_last_frame(itf):
        if log: logger.error(f'Given start frame number should be less than {get_last_frame(itf)} for the open file!')
        return None
    n_frs_updated = itf.DeleteFrames(start_frame, num_frames)
    return n_frs_updated

def update_marker_pos(itf, mkr_name, mkr_coords, start_frame=None, log=False):
    """
    Set the coordinates of a marker partially.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    mkr_coords : numpy array
        Marker coordinates.
    start_frame : int
        Frame number where setting will start.        
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, None, log)
    if not fr_check:
        if log: logger.error('Given "start_frame" is not proper!')
        return False
    n_frs = end_fr-start_fr+1
    if mkr_coords.ndim != 2 or mkr_coords.shape[0] != n_frs:
        if log: logger.error('The dimension of the input is not compatible!')
        return False    
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return False
    mkr_scale = get_marker_scale(itf)
    is_c3d_float = mkr_scale < 0
    is_c3d_float2 = [False, True][itf.GetDataType()-1]
    if is_c3d_float != is_c3d_float2:
        if log: logger.debug('C3D data type is determined by the POINT:SCALE parameter.')
    mkr_dtype = [np.int16, np.float32][is_c3d_float]
    scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
    if is_c3d_float:
        mkr_coords_unscaled = np.asarray(np.nan_to_num(mkr_coords), dtype=mkr_dtype)
    else:
        mkr_coords_unscaled = np.asarray(np.round(np.nan_to_num(mkr_coords)/scale_size), dtype=mkr_dtype)
    dtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_c3d_float]
    dtype_arr = pythoncom.VT_ARRAY|dtype
    for i in range(3):
        variant = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i])
        ret = itf.SetPointDataEx(mkr_idx, i, start_fr, variant)
    var_const = win32.VARIANT(dtype, 1)
    for i in range(3):
        for idx, val in enumerate(mkr_coords_unscaled[:,i]):
            if val == 1:
                ret = itf.SetPointData(mkr_idx, i, start_fr+idx, var_const)
    return [False, True][ret]
    
def update_marker_resid(itf, mkr_name, mkr_resid, start_frame=None, log=False):
    """
    Set the residual of a marker partially.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    mkr_resid : numpy array
        Marker residuals.
    start_frame : int
        Frame number where setting will start.            
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, None, log)
    if not fr_check: 
        if log: logger.error('Given "start_frame" is not proper!')
        return False
    n_frs = end_fr-start_fr+1
    if mkr_resid.ndim != 1 or mkr_resid.shape[0] != n_frs:
        if log: logger.error('The dimension of the input is not compatible!')
        return False
    mkr_idx = get_marker_index(itf, mkr_name, log)
    if mkr_idx == -1 or mkr_idx is None: return False
    dtype = pythoncom.VT_R4
    dtype_arr = pythoncom.VT_ARRAY|dtype
    variant = win32.VARIANT(dtype_arr, mkr_resid)
    ret = itf.SetPointDataEx(mkr_idx, 3, start_fr, variant)
    var_const = win32.VARIANT(dtype, 1)
    for idx, val in enumerate(mkr_resid):
        if val == 1:
            ret = itf.SetPointData(mkr_idx, 3, start_fr+idx, var_const) 
    return [False, True][ret]

def recover_marker_rel(itf, tgt_mkr_name, cl_mkr_names, log=False):
    """
    Recover the trajectory of a marker using the relation between a group (cluster) of markers.
    
    The number of cluster marker names is fixed as 3.
    This function extrapolates the target marker coordinates for the frames where the cluster markers are available.
    
    First cluster marker (cl_mkr_names[0]) will be used as the origin of the LCS(Local Coordinate System).
    Second cluster marker (cl_mkr_names[1]) will be used in order to determine the X axis of the LCS.
    Third cluster marker (cl_mkr_names[2]) will be used in order to determine the XY plane of the LCS.    

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_mkr_name : str
        Target marker name.
    cl_mkr_names : list or tuple
        Cluster (group) marker names.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int
        Number of valid frames in the target marker after this function.
        
    Notes
    -----
    This function is adapted from 'recover_marker_rel()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    if log: logger.debug(f'Start recovery of {tgt_mkr_name} ...')
    n_total_frs = get_num_frames(itf)
    tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
    tgt_mkr_coords = tgt_mkr_data[:,0:3]
    tgt_mkr_resid = tgt_mkr_data[:,3]
    tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
    n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
    if n_tgt_mkr_valid_frs == 0:
        if log: logger.info(f'Recovery of {tgt_mkr_name} skipped: no valid target marker frame!')
        return False, n_tgt_mkr_valid_frs
    if n_tgt_mkr_valid_frs == n_total_frs:
        if log: logger.info(f'Recovery of {tgt_mkr_name} skipped: all target marker frames valid!')
        return False, n_tgt_mkr_valid_frs
    dict_cl_mkr_coords = {}
    dict_cl_mkr_valid = {}
    cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
    for mkr in cl_mkr_names:
        mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
        dict_cl_mkr_coords[mkr] = mkr_data[:, 0:3]
        dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
        cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
    all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
    if not np.any(all_mkr_valid_mask):
        if log: logger.info(f'Recovery of {tgt_mkr_name} skipped: no common valid frame among markers!')
        return False, n_tgt_mkr_valid_frs
    cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
    if not np.any(cl_mkr_only_valid_mask):
        if log: logger.info(f'Recovery of {tgt_mkr_name} skipped: cluster markers not helpful!')
        return False, n_tgt_mkr_valid_frs
    all_mkr_valid_frs = np.where(all_mkr_valid_mask)[0]
    cl_mkr_only_valid_frs = np.where(cl_mkr_only_valid_mask)[0]
    p0 = dict_cl_mkr_coords[cl_mkr_names[0]]
    p1 = dict_cl_mkr_coords[cl_mkr_names[1]]
    p2 = dict_cl_mkr_coords[cl_mkr_names[2]] 
    vec0 = p1-p0
    vec1 = p2-p0
    vec0_norm = np.linalg.norm(vec0, axis=1, keepdims=True)
    vec1_norm = np.linalg.norm(vec1, axis=1, keepdims=True)
    vec0_unit = np.divide(vec0, vec0_norm, where=(vec0_norm!=0))
    vec1_unit = np.divide(vec1, vec1_norm, where=(vec1_norm!=0))
    vec2 = np.cross(vec0_unit, vec1_unit)
    vec2_norm = np.linalg.norm(vec2, axis=1, keepdims=True)
    vec2_unit = np.divide(vec2, vec2_norm, where=(vec2_norm!=0))
    vec_z = vec2_unit
    vec_x = vec0_unit
    vec_y = np.cross(vec_z, vec_x)
    mat_rot = np.array([vec_x.T, vec_y.T, vec_z.T]).T
    tgt_mkr_coords_rel = np.einsum('ij,ijk->ik', (tgt_mkr_coords-p0)[all_mkr_valid_mask], mat_rot[all_mkr_valid_mask])
    tgt_mkr_coords_recovered = np.zeros((cl_mkr_only_valid_frs.size, 3), dtype=np.float32)
    for idx, fr in np.ndenumerate(cl_mkr_only_valid_frs):
        search_idx = np.searchsorted(all_mkr_valid_frs, fr)
        if search_idx>=all_mkr_valid_frs.shape[0] or search_idx==0:
            tgt_coords_rel_idx = (np.abs(all_mkr_valid_frs-fr)).argmin()
            tgt_coords_rel = tgt_mkr_coords_rel[tgt_coords_rel_idx]
        else:
            idx1 = search_idx
            idx0 = search_idx-1
            fr1 = all_mkr_valid_frs[idx1]
            fr0 = all_mkr_valid_frs[idx0]
            a = np.float32(fr-fr0)
            b = np.float32(fr1-fr)
            tgt_coords_rel = (b*tgt_mkr_coords_rel[idx0]+a*tgt_mkr_coords_rel[idx1])/(a+b)
        tgt_mkr_coords_recovered[idx] = p0[fr]+np.dot(mat_rot[fr], tgt_coords_rel)
    tgt_mkr_coords[cl_mkr_only_valid_mask] = tgt_mkr_coords_recovered
    tgt_mkr_resid[cl_mkr_only_valid_mask] = 0.0
    update_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, log=log)
    update_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, log=log)
    n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
    if log: logger.info(f'Recovery of {tgt_mkr_name} is finished.')
    return True, n_tgt_mkr_valid_frs_updated

def recover_marker_rbt(itf, tgt_mkr_name, cl_mkr_names, log=False):
    """
    Recover the trajectory of a marker by rbt(rigid body transformation) using a group (cluster) markers.
    
    The number of cluster marker names is fixed as 3.
    This function extrapolates the target marker coordinates for the frames where the cluster markers are available.
    The order of the cluster markers will be sorted according to their relative distances from the target marker.    

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_mkr_name : str
        Target marker name.
    cl_mkr_names : list or tuple
        Cluster (group) marker names.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int
        Number of valid frames in the target marker after this function.

    Notes
    -----
    This function is adapted from 'recover_marker_rbt()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    if log: logger.debug(f'Start recovery of {tgt_mkr_name} ...')
    n_total_frs = get_num_frames(itf)
    tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
    tgt_mkr_coords = tgt_mkr_data[:,0:3]
    tgt_mkr_resid = tgt_mkr_data[:,3]
    tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
    n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
    if n_tgt_mkr_valid_frs == 0:
        if log: logger.info(f'Recovery of {tgt_mkr_name} skipped: no valid target marker frame!')
        return False, n_tgt_mkr_valid_frs
    if n_tgt_mkr_valid_frs == n_total_frs:
        if log: logger.info('Recovery of {tgt_mkr_name} skipped: all target marker frames valid!')
        return False, n_tgt_mkr_valid_frs    
    dict_cl_mkr_coords = {}
    dict_cl_mkr_valid = {}
    cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
    for mkr in cl_mkr_names:
        mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
        dict_cl_mkr_coords[mkr] = mkr_data[:,0:3]
        dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
        cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
    all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
    if not np.any(all_mkr_valid_mask):
        if log: logger.info('Recovery of {tgt_mkr_name} skipped: no common valid frame among markers!')
        return False, n_tgt_mkr_valid_frs
    cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
    if not np.any(cl_mkr_only_valid_mask):
        if log: logger.info('Recovery of {tgt_mkr_name} skipped: cluster markers not helpful!')
        return False, n_tgt_mkr_valid_frs
    all_mkr_valid_frs = np.where(all_mkr_valid_mask)[0]
    cl_mkr_only_valid_frs = np.where(cl_mkr_only_valid_mask)[0]
    dict_cl_mkr_dist = {}
    for mkr_name in cl_mkr_names:
        vec_diff = dict_cl_mkr_coords[mkr_name]-tgt_mkr_coords
        dict_cl_mkr_dist.update({mkr_name: np.nanmean(np.linalg.norm(vec_diff, axis=1))})
    cl_mkr_dist_sorted = sorted(dict_cl_mkr_dist.items(), key=lambda kv: kv[1])
    p0 = dict_cl_mkr_coords[cl_mkr_dist_sorted[0][0]]
    p1 = dict_cl_mkr_coords[cl_mkr_dist_sorted[1][0]]
    p2 = dict_cl_mkr_coords[cl_mkr_dist_sorted[2][0]]
    p3 = tgt_mkr_coords
    vec0 = p1-p0
    vec1 = p2-p0
    vec0_norm = np.linalg.norm(vec0, axis=1, keepdims=True)
    vec1_norm = np.linalg.norm(vec1, axis=1, keepdims=True)
    vec0_unit = np.divide(vec0, vec0_norm, where=(vec0_norm!=0))
    vec1_unit = np.divide(vec1, vec1_norm, where=(vec1_norm!=0))
    vec2 = np.cross(vec0_unit, vec1_unit)
    vec2_norm = np.linalg.norm(vec2, axis=1, keepdims=True)
    vec2_unit = np.divide(vec2, vec2_norm, where=(vec2_norm!=0))
    vec3 = p3-p0
    vec_z = vec2_unit
    vec_x = vec0_unit
    vec_y = np.cross(vec_z, vec_x)
    mat_rot = np.array([vec_x.T, vec_y.T, vec_z.T]).T
    for idx, fr in np.ndenumerate(cl_mkr_only_valid_frs):
        search_idx = np.searchsorted(all_mkr_valid_frs, fr)
        if search_idx == 0:
            fr0 = all_mkr_valid_frs[0]
            rot_fr0_to_fr = np.dot(mat_rot[fr], mat_rot[fr0].T)
            vt_fr0 = np.dot(rot_fr0_to_fr, vec3[fr0])
            vc = vt_fr0
        elif search_idx >= all_mkr_valid_frs.shape[0]:
            fr1 = all_mkr_valid_frs[all_mkr_valid_frs.shape[0]-1]
            rot_fr1_to_fr = np.dot(mat_rot[fr], mat_rot[fr1].T)
            vt_fr1 = np.dot(rot_fr1_to_fr, vec3[fr1])
            vc = vt_fr1
        else:
            fr0 = all_mkr_valid_frs[search_idx-1]
            fr1 = all_mkr_valid_frs[search_idx]
            rot_fr0_to_fr = np.dot(mat_rot[fr], mat_rot[fr0].T)
            rot_fr1_to_fr = np.dot(mat_rot[fr], mat_rot[fr1].T)
            vt_fr0 = np.dot(rot_fr0_to_fr, vec3[fr0])
            vt_fr1 = np.dot(rot_fr1_to_fr, vec3[fr1])
            a = np.float32(fr-fr0)
            b = np.float32(fr1-fr)
            vc = (b*vt_fr0+a*vt_fr1)/(a+b)
        tgt_mkr_coords[fr] = p0[fr]+vc
        tgt_mkr_resid[fr] = 0.0
    update_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, log=log)
    update_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, log=log)
    n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
    if log: logger.info(f'Recovery of {tgt_mkr_name} is finished.')
    return True, n_tgt_mkr_valid_frs_updated

def fill_marker_gap_rbt(itf, tgt_mkr_name, cl_mkr_names, log=False):
    """
    Fill the gaps in the trajectory of a marker by rbt(rigid body transformation) using a group (cluster) markers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_mkr_name : str
        Target marker name.
    cl_mkr_names : list or tuple
        Cluster (group) marker names.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int
        Number of valid frames in the target marker after this function.

    Notes
    -----
    This function is adapted from 'fill_marker_gap_rbt()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    def RBT(A, B):
        Ac = A.mean(axis=0)
        Bc = B.mean(axis=0)
        C = np.dot((B-Bc).T, (A-Ac))
        U, S, Vt = np.linalg.svd(C)
        R = np.dot(U, np.dot(np.diag([1, 1, np.linalg.det(np.dot(U, Vt))]), Vt))
        t = Bc-np.dot(R, Ac)
        err_vec = np.dot(R, A.T).T+t-B
        err_norm = np.linalg.norm(err_vec, axis=1)
        mean_err_norm = np.mean(err_norm)
        return R, t, err_vec, err_norm, mean_err_norm
    if log: logger.debug(f'Start gap filling of {tgt_mkr_name} ...')     
    n_total_frs = get_num_frames(itf)
    tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
    tgt_mkr_coords = tgt_mkr_data[:,0:3]
    tgt_mkr_resid = tgt_mkr_data[:,3]
    tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
    n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
    if n_tgt_mkr_valid_frs == 0:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no valid target marker frame!')
        return False, n_tgt_mkr_valid_frs
    if n_tgt_mkr_valid_frs == n_total_frs:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: all target marker frames valid!')
        return False , n_tgt_mkr_valid_frs   
    dict_cl_mkr_coords = {}
    dict_cl_mkr_valid = {}
    cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
    for mkr in cl_mkr_names:
        mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
        dict_cl_mkr_coords[mkr] = mkr_data[:,0:3]
        dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
        cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
    all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
    if not np.any(all_mkr_valid_mask):
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no common valid frame among markers!')
        return False, n_tgt_mkr_valid_frs
    cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
    if not np.any(cl_mkr_only_valid_mask):
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: cluster markers not helpful!')
        return False, n_tgt_mkr_valid_frs
    all_mkr_valid_frs = np.where(all_mkr_valid_mask)[0]
    cl_mkr_only_valid_frs = np.where(cl_mkr_only_valid_mask)[0]
    b_updated = False
    for idx, fr in np.ndenumerate(cl_mkr_only_valid_frs):
        search_idx = np.searchsorted(all_mkr_valid_frs, fr)
        if search_idx == 0:
            fr0 = all_mkr_valid_frs[0]
            fr1 = all_mkr_valid_frs[1]
        elif search_idx >= all_mkr_valid_frs.shape[0]:
            fr0 = all_mkr_valid_frs[all_mkr_valid_frs.shape[0]-2]
            fr1 = all_mkr_valid_frs[all_mkr_valid_frs.shape[0]-1]
        else:
            fr0 = all_mkr_valid_frs[search_idx-1]
            fr1 = all_mkr_valid_frs[search_idx]
        if fr <= fr0 or fr >= fr1: continue
        if ~cl_mkr_valid_mask[fr0] or ~cl_mkr_valid_mask[fr1]: continue
        if np.any(~cl_mkr_valid_mask[fr0:fr1+1]): continue
        cl_mkr_coords_fr0 = np.zeros((len(cl_mkr_names), 3), dtype=np.float32)
        cl_mkr_coords_fr1 = np.zeros((len(cl_mkr_names), 3), dtype=np.float32)
        cl_mkr_coords_fr = np.zeros((len(cl_mkr_names), 3), dtype=np.float32)
        for cnt, mkr in enumerate(cl_mkr_names):
            cl_mkr_coords_fr0[cnt,:] = dict_cl_mkr_coords[mkr][fr0,:]
            cl_mkr_coords_fr1[cnt,:] = dict_cl_mkr_coords[mkr][fr1,:]
            cl_mkr_coords_fr[cnt,:] = dict_cl_mkr_coords[mkr][fr,:]
        rot_fr0, trans_fr0, _, _, _ = RBT(cl_mkr_coords_fr0, cl_mkr_coords_fr)
        rot_fr1, trans_fr1, _, _, _ = RBT(cl_mkr_coords_fr1, cl_mkr_coords_fr)
        tgt_mkr_coords_fr_fr0 = np.dot(rot_fr0, tgt_mkr_coords[fr0])+trans_fr0
        tgt_mkr_coords_fr_fr1 = np.dot(rot_fr1, tgt_mkr_coords[fr1])+trans_fr1
        tgt_mkr_coords[fr] = (tgt_mkr_coords_fr_fr1-tgt_mkr_coords_fr_fr0)*np.float32(fr-fr0)/np.float32(fr1-fr0)+tgt_mkr_coords_fr_fr0
        tgt_mkr_resid[fr] = 0.0        
        b_updated = True        
    if b_updated:
        update_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, log=log)
        update_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, log=log)
        n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is finished.')
        return True, n_tgt_mkr_valid_frs_updated
    else:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is skipped.')
        return False, n_tgt_mkr_valid_frs

def fill_marker_gap_pattern(itf, tgt_mkr_name, dnr_mkr_name, log=False):
    """
    Fill the gaps in a given target marker coordinates using the donor marker coordinates by linear interpolation.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_mkr_name : str
        Target marker name.
    dnr_mkr_name : str
        Donor marker name.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int
        Number of valid frames in the target marker after this function.

    Notes
    -----
    This function is adapted from 'fill_marker_gap_pattern2()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    if log: logger.debug(f'Start gap filling of {tgt_mkr_name} ...')    
    n_total_frs = get_num_frames(itf)
    tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
    tgt_mkr_coords = tgt_mkr_data[:, 0:3]
    tgt_mkr_resid = tgt_mkr_data[:, 3]
    tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
    n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
    if n_tgt_mkr_valid_frs == 0:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no valid target marker frame!')
        return False, n_tgt_mkr_valid_frs
    if n_tgt_mkr_valid_frs == n_total_frs:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: all target marker frames valid!')
        return False , n_tgt_mkr_valid_frs    
    dnr_mkr_data = get_marker_data(itf, dnr_mkr_name, blocked_nan=False, log=log)
    dnr_mkr_coords = dnr_mkr_data[:, 0:3]
    dnr_mkr_resid = dnr_mkr_data[:, 3]
    dnr_mkr_valid_mask = np.where(np.isclose(dnr_mkr_resid, -1), False, True)
    if not np.any(dnr_mkr_valid_mask):
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no valid donor marker frame!')
        return False, n_tgt_mkr_valid_frs    
    both_mkr_valid_mask = np.logical_and(tgt_mkr_valid_mask, dnr_mkr_valid_mask)
    if not np.any(both_mkr_valid_mask):
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no valid common frame between target and donor markers!')
        return False, n_tgt_mkr_valid_frs        
    b_updated = False
    tgt_mkr_invalid_frs = np.where(~tgt_mkr_valid_mask)[0]
    both_mkr_valid_frs = np.where(both_mkr_valid_mask)[0]
    for idx, fr in np.ndenumerate(tgt_mkr_invalid_frs):
        search_idx = np.searchsorted(both_mkr_valid_frs, fr)
        if search_idx == 0:
            fr0 = both_mkr_valid_frs[0]
            fr1 = both_mkr_valid_frs[1]
        elif search_idx >= both_mkr_valid_frs.shape[0]:
            fr0 = both_mkr_valid_frs[both_mkr_valid_frs.shape[0]-2]
            fr1 = both_mkr_valid_frs[both_mkr_valid_frs.shape[0]-1]
        else:
            fr0 = both_mkr_valid_frs[search_idx-1]
            fr1 = both_mkr_valid_frs[search_idx]
        if fr <= fr0 or fr >= fr1: continue
        if ~dnr_mkr_valid_mask[fr0] or ~dnr_mkr_valid_mask[fr1]: continue
        if np.any(~dnr_mkr_valid_mask[fr0:fr1+1]): continue    
        v_tgt = (tgt_mkr_coords[fr1]-tgt_mkr_coords[fr0])*np.float32(fr-fr0)/np.float32(fr1-fr0)+tgt_mkr_coords[fr0]
        v_dnr = (dnr_mkr_coords[fr1]-dnr_mkr_coords[fr0])*np.float32(fr-fr0)/np.float32(fr1-fr0)+dnr_mkr_coords[fr0]
        new_coords = v_tgt-v_dnr+dnr_mkr_coords[fr]      
        tgt_mkr_coords[fr] = new_coords
        tgt_mkr_resid[fr] = 0.0        
        b_updated = True
    if b_updated:
        update_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, log=log)
        update_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, log=log)
        n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is finished.')
        return True, n_tgt_mkr_valid_frs_updated
    else:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is skipped.')
        return False, n_tgt_mkr_valid_frs

def fill_marker_gap_interp(itf, tgt_mkr_name, k=3, search_span_offset=5, min_needed_frs=10, log=False):
    """
    Fill the gaps in a given target marker coordinates using scipy.interpolate.InterpolatedUnivariateSpline function.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    tgt_mkr_name : str
        Target marker name.
    k : int, optional
        Degrees of smoothing spline. The default is 3.
    search_span_offset : int, optional
        Offset for backward and forward search spans. The default is 5.
    min_needed_frs : int, optional
        Minimum required valid frames in a search span. The default is 10.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.
    int
        Number of valid frames in the target marker after this function.

    Notes
    -----
    This function is adapted from 'fill_marker_gap_interp()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    if log: logger.debug(f'Start gap filling of {tgt_mkr_name} ...')
    n_total_frs = get_num_frames(itf)
    tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
    tgt_mkr_coords = tgt_mkr_data[:, 0:3]
    tgt_mkr_resid = tgt_mkr_data[:, 3]
    tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
    n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)    
    if n_tgt_mkr_valid_frs == 0:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: no valid target marker frame!')
        return False, n_tgt_mkr_valid_frs
    if n_tgt_mkr_valid_frs == n_total_frs:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} skipped: all target marker frames valid!')
        return False , n_tgt_mkr_valid_frs     
    b_updated = False
    tgt_mkr_invalid_frs = np.where(~tgt_mkr_valid_mask)[0]
    tgt_mkr_invalid_gaps = np.split(tgt_mkr_invalid_frs, np.where(np.diff(tgt_mkr_invalid_frs)!=1)[0]+1)
    for gap in tgt_mkr_invalid_gaps:
        if gap.size == 0: continue
        if gap.min()==0 or gap.max()==n_total_frs-1: continue
        search_span = np.int(np.ceil(gap.size/2))+search_span_offset
        itpl_cand_frs_mask = np.zeros((n_total_frs,), dtype=bool)
        for i in range(gap.min()-1, gap.min()-1-search_span, -1):
            if i>=0: itpl_cand_frs_mask[i]=True
        for i in range(gap.max()+1, gap.max()+1+search_span, 1):
            if i<n_total_frs: itpl_cand_frs_mask[i]=True
        itpl_cand_frs_mask = np.logical_and(itpl_cand_frs_mask, tgt_mkr_valid_mask)
        if np.sum(itpl_cand_frs_mask) < min_needed_frs: continue
        itpl_cand_frs = np.where(itpl_cand_frs_mask)[0]
        itpl_cand_coords = tgt_mkr_coords[itpl_cand_frs, :]
        fun_itpl_x = InterpolatedUnivariateSpline(itpl_cand_frs, itpl_cand_coords[:,0], k=k, ext='const')
        fun_itpl_y = InterpolatedUnivariateSpline(itpl_cand_frs, itpl_cand_coords[:,1], k=k, ext='const')
        fun_itpl_z = InterpolatedUnivariateSpline(itpl_cand_frs, itpl_cand_coords[:,2], k=k, ext='const')
        itpl_x = fun_itpl_x(gap)
        itpl_y = fun_itpl_y(gap)
        itpl_z = fun_itpl_z(gap)
        for idx, fr in enumerate(gap):
            tgt_mkr_coords[fr,0] = itpl_x[idx]
            tgt_mkr_coords[fr,1] = itpl_y[idx]
            tgt_mkr_coords[fr,2] = itpl_z[idx]
            tgt_mkr_resid[fr] = 0.0        
        b_updated = True            
    if b_updated:
        update_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, log=log)
        update_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, log=log)
        n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is finished.')
        return True, n_tgt_mkr_valid_frs_updated
    else:
        if log: logger.info(f'Gap filling of {tgt_mkr_name} is skipped.')
        return False, n_tgt_mkr_valid_frs
    