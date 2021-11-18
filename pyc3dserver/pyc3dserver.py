"""
MIT License

Copyright (c) 2020 Moon Ki Jung, Dario Farina

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

__author__ = 'Moon Ki Jung, Dario Farina'
__version__ = '0.2.0'

import os
import pythoncom
import win32com.client as win32
import math
import numpy as np
from scipy.signal import butter, filtfilt
from scipy.interpolate import InterpolatedUnivariateSpline
import re
import logging

logger_name = 'pyc3dserver'
logger = logging.getLogger(logger_name)
logger.setLevel('CRITICAL')
logger.addHandler(logging.NullHandler())

def filt_bw_bp(data, fc_low, fc_high, fs, order=2):
    nyq = 0.5 * fs
    low = fc_low / nyq
    high = fc_high / nyq
    b, a = butter(order, [low, high], analog=False, btype='bandpass', output='ba')
    axis = -1 if len(data.shape)==1 else 0
    y = filtfilt(b, a, data, axis, padtype='odd', padlen=3*(max(len(b),len(a))-1))
    return y

def filt_bw_bs(data, fc_low, fc_high, fs, order=2):
    nyq = 0.5 * fs
    low = fc_low / nyq
    high = fc_high / nyq
    b, a = butter(order, [low, high], analog=False, btype='bandstop', output='ba')
    axis = -1 if len(data.shape)==1 else 0
    y = filtfilt(b, a, data, axis, padtype='odd', padlen=3*(max(len(b),len(a))-1))
    return y

def filt_bw_lp(data, fc_low, fs, order=2):
    nyq = 0.5 * fs
    low = fc_low / nyq
    b, a = butter(order, low, analog=False, btype='lowpass', output='ba')
    axis = -1 if len(data.shape)==1 else 0
    y = filtfilt(b, a, data, axis, padtype='odd', padlen=3*(max(len(b),len(a))-1))
    return y

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
    except pythoncom.com_error as err:
        if log: logger.error(err.strerror)
        raise

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
    try:
        if log: logger.debug(f'Opening the file: "{f_path}"')
        if not os.path.exists(f_path):
            err_msg = 'File path does not exist'
            raise FileNotFoundError(err_msg)
        ret = itf.Open(f_path, 3)
        if strict_param_check:
            itf.SetStrictParameterChecking(1)
        else:
            itf.SetStrictParameterChecking(0)
        if ret == 0:
            if log: logger.info(f'File is opened')
            return True
        else:
            err_msg = f'File can not be opened'
            raise RuntimeError(err_msg)         
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except FileNotFoundError as err:
        if log: logger.error(err)
        raise        
    except RuntimeError as err:
        if log: logger.error(err)
        raise        

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
    try:
        if log: logger.debug(f'Saving the file: "{f_path}"')
        if compress_param_blocks:
            itf.CompressParameterBlocks(1)
        else:
            itf.CompressParameterBlocks(0)
        ret = itf.SaveFile(f_path, f_type)
        if ret == 1:
            if log: logger.info(f'File is saved')
            return True
        else:
            err_msg = f'File can not be saved'
            raise RuntimeError(err_msg)        
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise

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
    if log: logger.info(f'File is closed')
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
    try:
        dict_f_type = {1:'INTEL', 2:'DEC', 3:'SGI'}
        f_type = itf.GetFileType()
        return dict_f_type.get(f_type, None)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        dict_data_type = {1:'INTEGER', 2:'REAL'}
        data_type = itf.GetDataType()
        return dict_data_type.get(data_type, None)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
        return np.int32(first_fr)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
        return np.int32(last_fr)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)
        n_frs = last_fr-first_fr+1
        return np.int32(n_frs)        
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)
        if start_frame is None:
            start_fr = first_fr
        else:
            if start_frame < first_fr:
                err_msg = f'"start_frame" should be equal or greater than {first_fr}'
                raise ValueError(err_msg)
            start_fr = start_frame
        if end_frame is None:
            end_fr = last_fr
        else:
            if end_frame > last_fr:
                err_msg = f'"end_frame" should be equal or less than {last_fr}'
                raise ValueError(err_msg)
            end_fr = end_frame
        if not (start_fr <= end_fr):
            err_msg = f'"end_frame" should be greater than "start_frame"'
            raise ValueError(err_msg)
        return True, start_fr, end_fr
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise

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
        return np.float32(vid_fps)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
        return np.int32(av_ratio)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        vid_fps = get_video_fps(itf, log=log)
        av_ratio = get_analog_video_ratio(itf, log=log)
        return np.float32(vid_fps*np.float32(av_ratio))
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

def get_video_frames(itf, log=False):
    """
    Return an integer-type numpy array that contains the video frame numbers between the start and the end frames.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    frs : numpy array
        An integer-type numpy array of the video frame numbers.

    """
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)
        n_frs = last_fr-first_fr+1
        frs = np.linspace(start=first_fr, stop=last_fr, num=n_frs, dtype=np.int32)
        return frs       
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
        
def get_analog_frames(itf, log=False):
    """
    Return a float-type numpy array that contains the analog frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    frs : numpy array
        A float-type numpy array of the analog frame numbers.

    """
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)  
        av_ratio = get_analog_video_ratio(itf, log=log)
        start_fr = np.float32(first_fr)
        end_fr = np.float32(last_fr)+np.float32(av_ratio-1)/np.float32(av_ratio)
        n_frs = last_fr-first_fr+1
        analog_steps = n_frs*av_ratio
        frs = np.linspace(start=start_fr, stop=end_fr, num=analog_steps, dtype=np.float32)
        return frs
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

def get_video_times(itf, from_zero=True, log=False):
    """
    Return a float-type numpy array that contains the times corresponding to the video frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    from_zero : bool, optional
        Whether the return time array should start from zero or not. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    t : numpy array
        A float-type numpy array of the times corresponding to the video frame numbers.

    """
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)
        vid_fps = get_video_fps(itf, log=log)
        offset_fr = first_fr if from_zero else 0
        start_t = np.float32(first_fr-offset_fr)/vid_fps
        end_t = np.float32(last_fr-offset_fr)/vid_fps
        n_frs = last_fr-first_fr+1
        t = np.linspace(start=start_t, stop=end_t, num=n_frs, dtype=np.float32)
        return t        
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

def get_analog_times(itf, from_zero=True, log=False):
    """
    Return a float-type array that contains the times corresponding to the analog frame numbers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    from_zero : bool, optional
        Whether the return time array should start from zero or not. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.
        
    Returns
    -------
    t : numpy array
        A float-type numpy array of the times corresponding to the analog frame numbers.

    """
    try:
        first_fr = get_first_frame(itf, log=log)
        last_fr = get_last_frame(itf, log=log)
        vid_fps = get_video_fps(itf, log=log)
        analog_fps = get_analog_fps(itf, log=log)
        av_ratio = get_analog_video_ratio(itf, log=log)
        offset_fr = first_fr if from_zero else 0
        start_t = np.float32(first_fr-offset_fr)/vid_fps
        end_t = np.float32(last_fr-offset_fr)/vid_fps+np.float32(av_ratio-1)/analog_fps
        vid_steps = last_fr-first_fr+1
        analog_steps = vid_steps*av_ratio
        t = np.linspace(start=start_t, stop=end_t, num=analog_steps, dtype=np.float32)
        return t        
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        mkr_names = []
        idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
        if idx_pt_labels == -1:
            if log: logger.warning('POINT:LABELS does not exist')
            return None
        n_pt_labels = itf.GetParameterLength(idx_pt_labels)
        if n_pt_labels < 1:
            if log: logger.warning('No item under POINT:LABELS')
            return None
        idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
        if idx_pt_used == -1:
            if log: logger.warning('POINT:USED does not exist')
            return None
        n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
        if n_pt_used < 1:
            if log: logger.warning('POINT:USED is zero')
            return None
        for i in range(n_pt_labels):
            if i < n_pt_used:
                mkr_names.append(itf.GetParameterValue(idx_pt_labels, i))
        return mkr_names
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
        if idx_pt_labels == -1:
            if log: logger.warning('POINT:LABELS does not exist')
            return None
        n_pt_labels = itf.GetParameterLength(idx_pt_labels)
        if n_pt_labels < 1:
            if log: logger.warning('No item under POINT:LABELS')
            return None
        idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
        if idx_pt_used == -1:
            if log: logger.warning('POINT:USED does not exist')
            return None
        n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
        if n_pt_used < 1:
            if log: logger.warning('POINT:USED is zero')
            return None
        mkr_idx = -1
        for i in range(n_pt_labels):
            if i < n_pt_used:
                tgt_name = itf.GetParameterValue(idx_pt_labels, i)
                if tgt_name == mkr_name:
                    mkr_idx = i
                    break  
        if mkr_idx == -1:
            if log: logger.warning(f'"{mkr_name}" does not exist')
        return mkr_idx
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        idx_pt_units = itf.GetParameterIndex('POINT', 'UNITS')
        if idx_pt_units == -1: 
            if log: logger.warning('POINT:UNITS does not exist')
            return None
        n_items = itf.GetParameterLength(idx_pt_units)
        if n_items != 1: 
            if log: logger.warning('No proper item under POINT:UNITS')
            return None
        # unit = itf.GetParameterValue(idx_pt_units, n_items-1)
        unit = itf.GetParameterValue(idx_pt_units, 0)
        return unit
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        idx_pt_scale = itf.GetParameterIndex('POINT', 'SCALE')
        if idx_pt_scale == -1:
            if log: logger.warning('POINT:SCALE does not exist')
            return None
        n_items = itf.GetParameterLength(idx_pt_scale)
        if n_items != 1:
            if log: logger.warning('No proper item under POINT:SCALE')
            return None
        # scale = np.float32(itf.GetParameterValue(idx_pt_scale, n_items-1))
        scale = np.float32(itf.GetParameterValue(idx_pt_scale, 0))
        return scale
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            if log: logger.warning(f'Unable to get the index of "{mkr_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        n_frs = end_fr-start_fr+1
        mkr_data = np.full((n_frs, 4), np.nan, dtype=np.float32)
        if start_fr == end_fr:
            for i in range(3):
                mkr_data[:,i] = np.asarray(itf.GetPointData(mkr_idx, i, start_fr, '1'), dtype=np.float32)
            mkr_data[:,3] = np.asarray(itf.GetPointResidual(mkr_idx, start_fr), dtype=np.float32)
        else:
            for i in range(3):
                mkr_data[:,i] = np.asarray(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '1'), dtype=np.float32)
            mkr_data[:,3] = np.asarray(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
        if blocked_nan:
            mkr_null_masks = np.where(np.isclose(mkr_data[:,3], -1), True, False)
            mkr_data[mkr_null_masks,0:3] = np.nan 
        return mkr_data
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            if log: logger.warning(f'Unable to get the index of "{mkr_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        n_frs = end_fr-start_fr+1
        mkr_scale = get_marker_scale(itf, log=log)
        if mkr_scale is None:
            if log: logger.warning(f'Unable to get the marker scale factor')
            return None
        is_c3d_float = mkr_scale < 0
        is_c3d_float2 = [False, True][itf.GetDataType()-1]
        if is_c3d_float != is_c3d_float2:
            if log: logger.debug(f'C3D data type is determined by POINT:SCALE')
        mkr_dtype = [[[np.int16, np.float32][is_c3d_float], np.float32][scaled], np.float32][blocked_nan]
        mkr_data = np.zeros((n_frs, 3), dtype=mkr_dtype)
        b_scaled = ['0', '1'][scaled]
        if start_fr == end_fr:
            for i in range(3):
                mkr_data[:,i] = np.asarray(itf.GetPointData(mkr_idx, i, start_fr, b_scaled), dtype=mkr_dtype)
        else:
            for i in range(3):
                mkr_data[:,i] = np.asarray(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, b_scaled), dtype=mkr_dtype)
        if blocked_nan:
            if start_fr == end_fr:
                mkr_resid = np.asarray(itf.GetPointResidual(mkr_idx, start_fr), dtype=np.float32)
            else:
                mkr_resid = np.asarray(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
            mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
            mkr_data[mkr_null_masks,:] = np.nan  
        return mkr_data
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            if log: logger.warning(f'Unable to get the index of "{mkr_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        n_frs = end_fr-start_fr+1
        mkr_scale = get_marker_scale(itf, log=log)
        if mkr_scale is None:
            if log: logger.warning(f'Unable to get the marker scale factor')
            return None        
        is_c3d_float = mkr_scale < 0
        is_c3d_float2 = [False, True][itf.GetDataType()-1]
        if is_c3d_float != is_c3d_float2:
            if log: logger.debug(f'C3D data type is determined by POINT:SCALE')
        mkr_dtype = [[[np.int16, np.float32][is_c3d_float], np.float32][scaled], np.float32][blocked_nan]
        mkr_data = np.zeros((n_frs, 3), dtype=mkr_dtype)
        scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
        if start_fr == end_fr:
            for i in range(3):
                if scaled:
                    mkr_data[:,i] = np.asarray(itf.GetPointData(mkr_idx, i, start_fr, '0'), dtype=mkr_dtype)*scale_size
                else:
                    mkr_data[:,i] = np.asarray(itf.GetPointData(mkr_idx, i, start_fr, '0'), dtype=mkr_dtype)                
        else:
            for i in range(3):
                if scaled:
                    mkr_data[:,i] = np.asarray(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '0'), dtype=mkr_dtype)*scale_size
                else:
                    mkr_data[:,i] = np.asarray(itf.GetPointDataEx(mkr_idx, i, start_fr, end_fr, '0'), dtype=mkr_dtype)
        if blocked_nan:
            if start_fr == end_fr:
                mkr_resid = np.asarray(itf.GetPointResidual(mkr_idx, start_fr), dtype=np.float32)
            else:
                mkr_resid = np.asarray(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
            mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
            mkr_data[mkr_null_masks,:] = np.nan            
        return mkr_data
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            if log: logger.warning(f'Unable to get the index of "{mkr_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        if start_fr == end_fr:
            mkr_resid = np.asarray(itf.GetPointResidual(mkr_idx, start_fr), dtype=np.float32)
        else:
            mkr_resid = np.asarray(itf.GetPointResidualEx(mkr_idx, start_fr, end_fr), dtype=np.float32)
        return mkr_resid
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        sig_names = []
        idx_anl_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        if idx_anl_labels == -1:
            if log: logger.warning('ANALOG:LABELS does not exist')
            return None
        n_anl_labels = itf.GetParameterLength(idx_anl_labels)
        if n_anl_labels < 1:
            if log: logger.warning('No item under ANALOG:LABELS')
            return None
        idx_anl_used = itf.GetParameterIndex('ANALOG', 'USED')
        if idx_anl_used == -1:
            if log: logger.warning('ANALOG:USED does not exist')
            return None        
        n_anl_used = itf.GetParameterValue(idx_anl_used, 0)    
        for i in range(n_anl_labels):
            if i < n_anl_used:
                sig_names.append(itf.GetParameterValue(idx_anl_labels, i))
        return sig_names
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        idx_anl_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        if idx_anl_labels == -1:
            if log: logger.warning('ANALOG:LABELS does not exist')
            return None
        n_anl_labels = itf.GetParameterLength(idx_anl_labels)
        if n_anl_labels < 1:
            if log: logger.warning('No item under ANALOG:LABELS')
            return None
        idx_anl_used = itf.GetParameterIndex('ANALOG', 'USED')
        if idx_anl_used == -1:
            if log: logger.warning('ANALOG:USED does not exist')
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
            if log: logger.warning(f'"{sig_name}" does not exist')
        return sig_idx
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        par_idx = itf.GetParameterIndex('ANALOG', 'GEN_SCALE')
        if par_idx == -1:
            if log: logger.warning('ANALOG:GEN_SCALE does not exist')
            return None
        n_items = itf.GetParameterLength(par_idx)
        if n_items != 1:
            if log: logger.warning('No proper item under ANALOG:GEN_SCALE')
            return None
        # gen_scale = np.float32(itf.GetParameterValue(par_idx, n_items-1))
        gen_scale = np.float32(itf.GetParameterValue(par_idx, 0))
        return gen_scale
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        par_idx = itf.GetParameterIndex('ANALOG', 'FORMAT')
        if par_idx == -1:
            if log: logger.debug('ANALOG:FORMAT does not exist')
            return None
        n_items = itf.GetParameterLength(par_idx)
        if n_items != 1:
            if log: logger.debug('No proper item under ANALOG:FORMAT')
            return None    
        # sig_format = itf.GetParameterValue(par_idx, n_items-1)
        sig_format = itf.GetParameterValue(par_idx, 0)
        return sig_format
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        par_idx = itf.GetParameterIndex('ANALOG', 'UNITS')
        if par_idx == -1:
            if log: logger.warning('ANALOG:UNITS does not exist')
            return None
        sig_unit = itf.GetParameterValue(par_idx, sig_idx)
        return sig_unit
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    
    
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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        par_idx = itf.GetParameterIndex('ANALOG', 'SCALE')
        if par_idx == -1:
            if log: logger.warning('ANALOG:SCALE does not exist')
            return None
        sig_scale = np.float32(itf.GetParameterValue(par_idx, sig_idx))
        return sig_scale
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    
    
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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        par_idx = itf.GetParameterIndex('ANALOG', 'OFFSET')
        if par_idx == -1:
            if log: logger.warning('ANALOG:OFFSET does not exist')
            return None
        sig_format = get_analog_format(itf, log=log)
        is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
        par_dtype = [np.int16, np.uint16][is_sig_unsigned]
        sig_offset = par_dtype(itf.GetParameterValue(par_idx, sig_idx))
        return sig_offset
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    
            
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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        sig_format = get_analog_format(itf, log=log)
        is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')        
        mkr_scale = get_marker_scale(itf, log=log)
        if mkr_scale is None:
            if log: logger.warning(f'Unable to get the marker scale factor')
            return None        
        is_c3d_float = mkr_scale < 0
        is_c3d_float2 = [False, True][itf.GetDataType()-1]
        if is_c3d_float != is_c3d_float2:
            if log: logger.debug(f'C3D data type is determined by POINT:SCALE')
        sig_dtype = [[np.int16, np.uint16][is_sig_unsigned], np.float32][is_c3d_float]
        if start_fr == end_fr:
            av_ratio = get_analog_video_ratio(itf)
            sig = np.zeros((av_ratio,), dtype=sig_dtype)
            for i in range(av_ratio):
                sig[i] = sig_dtype(itf.GetAnalogData(sig_idx, start_fr, i+1, '0', 0, 0, '0'))
        else:
            sig = np.asarray(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=sig_dtype)
        return sig
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        if start_fr == end_fr:
            av_ratio = get_analog_video_ratio(itf)
            sig = np.zeros((av_ratio,), dtype=np.float32)
            for i in range(av_ratio):
                sig[i] = np.float32(itf.GetAnalogData(sig_idx, start_fr, i+1, '1', 0, 0, '0'))
        else:
            sig = np.asarray(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '1', 0, 0, '0'), dtype=np.float32)
        return sig
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            if log: logger.warning(f'Unable to get the index of "{sig_name}"')
            return None
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            if log: logger.warning('No valid conditions for "start_frame" and "end_frame"')
            return None
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_scale = get_analog_scale(itf, sig_name, log=log)
        sig_offset = np.float32(get_analog_offset(itf, sig_name, log=log))
        if start_fr == end_fr:
            av_ratio = get_analog_video_ratio(itf)
            sig_data = np.zeros((av_ratio,), dtype=np.float32)
            for i in range(av_ratio):
                sig_data[i] = np.float32(itf.GetAnalogData(sig_idx, start_fr, i+1, '0', 0, 0, '0'))
        else:
            sig_data = np.asarray(itf.GetAnalogDataEx(sig_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)
        sig = (sig_data-sig_offset)*sig_scale*gen_scale
        return sig
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
        
def get_group_params(itf, grp_name, par_names, desc=False, log=False):
    """
    Return desired parameter values under a specific group.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    grp_name : str
        Target group name.
    par_names : list
        Target parameter names.
    desc : bool, optional
        Whether to include the descriptions of group parameters. The default is False.        
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict_info : dict
        Dictionary of the desired paramters under a specific group.

    Notes
    -----
    For multi-dimensional arrays, the return array will be reshaped in order to reverse its axes.
    This approach is especially useful for some FORCE_PLATFORM parameters such as FORCE_PLATFORM:CORNERS and FORCE_PLATFORM:ORIGIN.
    By reversing the axes of the arrays, the first dimension will indicate the index of the force plate.
    However, there is an exception, FORCE_PLATFORM:CAL_MATRIX. See the reference [1].
    
    References
    ----------
    .. [1] https://www.c3d.org/HTML/Documents/forceplatformcalmatrix.htm

    """
    try:
        dict_dtype = {-1:str, 1:np.int8, 2:np.int32, 4:np.float32}
        dict_info = {}
        for name in par_names:
            par_idx = itf.GetParameterIndex(grp_name, name)
            if par_idx == -1:
                if log: logger.warning(f'{grp_name}:{name} does not exist')
                continue
            par_name = itf.GetParameterName(par_idx)
            if desc:
                par_desc = itf.GetParameterDescription(par_idx)
            par_len = itf.GetParameterLength(par_idx)
            par_type = itf.GetParameterType(par_idx)
            data_type = dict_dtype.get(par_type, None)
            par_num_dim = itf.GetParameterNumberDim(par_idx)
            par_dim = [itf.GetParameterDimension(par_idx, j) for j in range(par_num_dim)]
            par_data = []
            # special handling for 'ANALOG:OFFSET' parameter
            if grp_name=='ANALOG' and par_name=='OFFSET':
                sig_format = get_analog_format(itf, log=log)
                is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
                pre_dtype = [np.int16, np.uint16][is_sig_unsigned]
                for j in range(par_len):
                    par_data.append(pre_dtype(itf.GetParameterValue(par_idx, j)))
            else:
                for j in range(par_len):
                    par_data.append(itf.GetParameterValue(par_idx, j))
            if par_type == -1:
                # if len(par_data) == 1:
                if par_num_dim <= 1:
                    par_val = data_type(par_data[0])
                else:
                    par_val = np.reshape(np.asarray(par_data, dtype=data_type), par_dim[::-1][:-1])
            else:
                # if par_num_dim==0 or (par_num_dim==1 and par_dim[0]==1):
                if par_num_dim == 0:
                    par_val = data_type(par_data[0])
                else:
                    par_val = np.reshape(np.asarray(par_data, dtype=data_type), par_dim[::-1])
            # special handling for 'FORCE_PLATFORM:CAL_MATRIX' parameter
            if grp_name=='FORCE_PLATFORM' and par_name=='CAL_MATRIX':
                par_val = np.transpose(par_val, (0,2,1))
            if desc:
                dict_info[par_name] = {}
                dict_info[par_name].update({'VAL': par_val})
                dict_info[par_name].update({'DESC': par_desc})
            else:
                dict_info[par_name] = par_val
        return dict_info
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

def get_dict_header(itf, log=False):
    """
    Return the summarization of the C3D header information.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.        

    Returns
    -------
    dict_header : dict
        Dictionary of the C3D header information.

    """
    try:
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
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

def get_dict_groups(itf, desc=False, tgt_grp_names=None, log=False):
    """
    Return the dictionary of the groups.

    All the values in the dictionary structure are numpy arrays except the values of single scalar.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    desc : bool, optional
        Whether to include the descriptions of group parameters. The default is False.        
    tgt_grp_names: list or tuple, optional
        Target group names to extract. The default is None.
    log : bool, optional
        Whether to write logs or not. The default is False.        
    
    Returns
    -------
    dict_grps : dict
        Dictionary of the C3D header information.
        
    Notes
    -----
    For multi-dimensional arrays, the return array will be reshaped in order to reverse its axes.
    This approach is especially useful for some FORCE_PLATFORM parameters such as FORCE_PLATFORM:CORNERS and FORCE_PLATFORM:ORIGIN.
    By reversing the axes of the arrays, the first dimension will indicate the index of the force plate.
    However, there is an exception, FORCE_PLATFORM:CAL_MATRIX. See the reference [1].
    
    References
    ----------
    .. [1] https://www.c3d.org/HTML/Documents/forceplatformcalmatrix.htm
    """
    try:
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
            if desc: 
                par_desc = itf.GetParameterDescription(i)
            par_len = itf.GetParameterLength(i)
            par_type = itf.GetParameterType(i)
            data_type = dict_dtype.get(par_type, None)
            par_num_dim = itf.GetParameterNumberDim(i)
            par_dim = [itf.GetParameterDimension(i, j) for j in range(par_num_dim)]
            par_data = []
            # special handling for 'ANALOG:OFFSET' parameter
            if grp_name=='ANALOG' and par_name=='OFFSET':
                sig_format = get_analog_format(itf, log=log)
                is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
                pre_dtype = [np.int16, np.uint16][is_sig_unsigned]
                for j in range(par_len):
                    par_data.append(pre_dtype(itf.GetParameterValue(i, j)))
            else:
                for j in range(par_len):
                    par_data.append(itf.GetParameterValue(i, j))
            if par_type == -1:
                # if len(par_data) == 1:
                if par_num_dim <= 1:
                    par_val = data_type(par_data[0])
                else:
                    par_val = np.reshape(np.asarray(par_data, dtype=data_type), par_dim[::-1][:-1])
            else:
                # if par_num_dim==0 or (par_num_dim==1 and par_dim[0]==1):
                if par_num_dim == 0:
                    par_val = data_type(par_data[0])
                else:
                    par_val = np.reshape(np.asarray(par_data, dtype=data_type), par_dim[::-1])
            # special handling for 'FORCE_PLATFORM:CAL_MATRIX' parameter
            if grp_name=='FORCE_PLATFORM' and par_name=='CAL_MATRIX':
                par_val = np.transpose(par_val, (0,2,1))                    
            if desc:
                dict_grps[grp_name][par_name] = {}
                dict_grps[grp_name][par_name].update({'VAL': par_val})
                dict_grps[grp_name][par_name].update({'DESC': par_desc})
            else:
                dict_grps[grp_name][par_name] = par_val
        return dict_grps
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    

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
    try:
        start_fr = get_first_frame(itf, log=log)
        end_fr = get_last_frame(itf, log=log) 
        n_frs = end_fr-start_fr+1
        idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
        if idx_pt_labels == -1:
            if log: logger.warning('POINT:LABELS does not exist')
            return None
        n_pt_labels = itf.GetParameterLength(idx_pt_labels)
        if n_pt_labels < 1:
            if log: logger.warning('No item under POINT:LABELS')
            return None
        idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
        if idx_pt_used == -1:
            if log: logger.warning('POINT:USED does not exsit')
            return None        
        n_pt_used = itf.GetParameterValue(idx_pt_used, 0)
        if n_pt_used < 1:
            if log: logger.warning('POINT:USED is zero')
            return None
        idx_pt_desc = itf.GetParameterIndex('POINT', 'DESCRIPTIONS')
        if idx_pt_desc == -1:
            if log: logger.warning('POINT:DESCRIPTIONS does not exist')
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
                    mkr_data[:,j] = np.asarray(itf.GetPointDataEx(i, j, start_fr, end_fr, '1'), dtype=np.float32)
                if blocked_nan or resid:
                    mkr_resid = np.asarray(itf.GetPointResidualEx(i, start_fr, end_fr), dtype=np.float32)
                if blocked_nan:
                    mkr_null_masks = np.where(np.isclose(mkr_resid, -1), True, False)
                    mkr_data[mkr_null_masks,:] = np.nan
                dict_pts['DATA']['POS'].update({mkr_name: mkr_data})
                if resid:
                    dict_pts['DATA']['RESID'].update({mkr_name: mkr_resid})
                if mask:
                    mkr_mask = np.asarray(itf.GetPointMaskEx(i, start_fr, end_fr), dtype=str)
                    dict_pts['DATA']['MASK'].update({mkr_name: mkr_mask})
                if desc:
                    if i < n_pt_desc:
                        mkr_descs.append(itf.GetParameterValue(idx_pt_desc, i))
                    else:
                        mkr_descs.append('')
        dict_pts.update({'LABELS': np.asarray(mkr_names, dtype=str)})
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
                dict_pts.update({'DESCRIPTIONS': np.asarray(mkr_descs, dtype=str)})
        if frame: dict_pts.update({'FRAME': get_video_frames(itf, log=log)})
        if time: dict_pts.update({'TIME': get_video_times(itf, log=log)})
        return dict_pts
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise

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
    try:
        start_fr = get_first_frame(itf, log=log)
        end_fr = get_last_frame(itf, log=log)
        n_force_chs = 0
        idx_force_chs = itf.GetParameterIndex('FORCE_PLATFORM', 'CHANNEL')
        if idx_force_chs == -1: 
            if log: logger.warning(f'FORCE_PLATFORM:CHANNEL does not exist')
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
            if log: logger.warning('ANALOG:LABELS does not exist')
            return None
        n_analog_labels = itf.GetParameterLength(idx_analog_labels)
        if n_analog_labels < 1:
            if log: logger.warning('No item under ANALOG:LABELS')
            return None    
        idx_analog_used = itf.GetParameterIndex('ANALOG', 'USED')
        if idx_analog_used == -1:
            if log: logger.warning('ANALOG:USED does not exist')
            return None
        n_analog_used = itf.GetParameterValue(idx_analog_used, 0)
        if n_analog_used < 1:
            if log: logger.warning(f'ANALOG:USED is zero')
            return None    
        idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
        if idx_analog_scale == -1:
            if log: logger.warning('ANALOG:SCALE does not exist')
            return None       
        idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
        if idx_analog_offset == -1:
            if log: logger.warning('ANALOG:OFFSET does not exist')
            return None
        idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
        if idx_analog_units == -1:
            if log: logger.warning('ANALOG:UNITS does not exist')
            n_analog_units = 0
        else:
            n_analog_units = itf.GetParameterLength(idx_analog_units)
        idx_analog_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
        if idx_analog_desc == -1:
            if log: logger.warning('ANALOG:DESCRIPTIONS does not exist')
            n_analog_desc = 0
        else:
            n_analog_desc = itf.GetParameterLength(idx_analog_desc)
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_format = get_analog_format(itf, log=log)
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
                sig_val = (np.asarray(itf.GetAnalogDataEx(i, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-sig_offset)*sig_scale*gen_scale
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
        dict_analogs.update({'LABELS': np.asarray(analog_names, dtype=str)})
        idx_analog_rate = itf.GetParameterIndex('ANALOG', 'RATE')
        if idx_analog_rate != -1:
            n_analog_rate = itf.GetParameterLength(idx_analog_rate)
            if n_analog_rate == 1:
                dict_analogs.update({'RATE': np.float32(itf.GetParameterValue(idx_analog_rate, 0))})
        if idx_analog_units != -1:
            dict_analogs.update({'UNITS': np.asarray(analog_units, dtype=str)})
        if desc:
            if idx_analog_desc != -1:
                dict_analogs.update({'DESCRIPTIONS': np.asarray(analog_descs, dtype=str)})
        if frame: dict_analogs.update({'FRAME': get_analog_frames(itf, log=log)})
        if time: dict_analogs.update({'TIME': get_analog_times(itf, log=log)})
        return dict_analogs
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    
    
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
    try:
        start_fr = get_first_frame(itf, log=log)
        end_fr = get_last_frame(itf, log=log)
        idx_fp_used = itf.GetParameterIndex('FORCE_PLATFORM', 'USED')
        if idx_fp_used == -1: 
            if log: logger.warning(f'FORCE_PLATFORM:USED does not exist')
            return None
        n_fp_used = itf.GetParameterValue(idx_fp_used, 0)
        if n_fp_used < 1:
            if log: logger.warning(f'FORCE_PLATFORM:USED is zero')
            return None
        idx_force_chs = itf.GetParameterIndex('FORCE_PLATFORM', 'CHANNEL')
        if idx_force_chs == -1: 
            if log: logger.warning(f'FORCE_PLATFORM:CHANNEL does not exist')
            return None
        idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        if idx_analog_labels == -1:
            if log: logger.warning('ANALOG:LABELS does not exist')
            return None       
        idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
        if idx_analog_scale == -1:
            if log: logger.warning('ANALOG:SCALE does not exist')
            return None     
        idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
        if idx_analog_offset == -1:
            if log: logger.warning('ANALOG:OFFSET does not exist')
            return None
        idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
        if idx_analog_units == -1:
            if log: logger.warning('ANALOG:UNITS does not exist')
            n_analog_units = 0
        else:
            n_analog_units = itf.GetParameterLength(idx_analog_units)
        idx_analog_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
        if idx_analog_desc == -1:
            if log: logger.warning('ANALOG:DESCRIPTIONS does not exist')
            n_analog_desc = 0
        else:
            n_analog_desc = itf.GetParameterLength(idx_analog_desc)
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_format = get_analog_format(itf, log=log)
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
            ch_val = (np.asarray(itf.GetAnalogDataEx(ch_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-ch_offset)*ch_scale*gen_scale
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
        dict_forces.update({'LABELS': np.asarray(force_names, dtype=str)})
        idx_analog_rate = itf.GetParameterIndex('ANALOG', 'RATE')
        if idx_analog_rate != -1:
            n_analog_rate = itf.GetParameterLength(idx_analog_rate)
            if n_analog_rate == 1:
                dict_forces.update({'RATE': np.float32(itf.GetParameterValue(idx_analog_rate, 0))})
        if idx_analog_units != -1:
            dict_forces.update({'UNITS': np.asarray(force_units, dtype=str)})
        if desc:
            if idx_analog_desc != -1:
                dict_forces.update({'DESCRIPTIONS': np.asarray(force_descs, dtype=str)})
        if frame: dict_forces.update({'FRAME': get_analog_frames(itf, log=log)})
        if time: dict_forces.update({'TIME': get_analog_times(itf, log=log)})
        return dict_forces
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
        
def get_fp_params(itf, log=False):
    """
    Return desired parameter values under the 'FORCE_PLATFORM' group.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict
        Dictionary of the desired paramters under the 'FORCE_PLATFORM' group.

    """
    try:
        grp_name = 'FORCE_PLATFORM'
        par_names = ['TYPE', 'USED', 'ORIGIN', 'CORNERS', 'CHANNEL', 'CAL_MATRIX']
        return get_group_params(itf, grp_name, par_names, log=log)
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
        
def get_fp_output(itf, threshold=0.0, filt_fc=None, filt_order=2, cop_nan_to_num=True, log=False):
    """
    Return forces, moments and COP(center of pressure) from the force plates.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    threshold : float, optional
        Threshold value of Fz (force plate local) to determine the frames where all forces and moments will be zero.        
    filt_fc: float or list or tuple, optional
        Cut-off frequency of zero-lag butterworth low-pass filters for each ananlog channel. The default is None.
    filt_order: int, optional
        Order of butterworth zero-lag low-pass filters for each analog channcel. The default is 2.        
    cop_nan_to_num : bool, optional
        Whether NaN values of COP will be converted to zero or not. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    dict
        Dictionary of the desired output from force plates including forces, moments, and COP.
        
    Notes
    -----
    The output forces and moments are expressed based on the MKS system of units.
    The unit of forces is N, whereas the unit of moments is Nm. The unit of COP is m.
    The supported force plate types are 1, 2, 3 and 4. See the reference [1] for more details.
    This function is an implementation of the following theoretical reference [2].   
    
    References
    ----------
    .. [1] https://www.c3d.org/HTML/Documents/theforceplatformgroup.htm
    .. [2] http://www.kwon3d.com/theory/grf/cop.html
    """    
    try:
        start_fr = get_first_frame(itf, log=log)
        end_fr = get_last_frame(itf, log=log)
        idx_fp_used = itf.GetParameterIndex('FORCE_PLATFORM', 'USED')
        if idx_fp_used == -1: 
            if log: logger.warning('FORCE_PLATFORM:USED does not exist')
            return None
        n_fp_used = itf.GetParameterValue(idx_fp_used, 0)
        if n_fp_used < 1:
            if log: logger.warning('FORCE_PLATFORM:USED is zero')
            return None
        idx_fp_chs = itf.GetParameterIndex('FORCE_PLATFORM', 'CHANNEL')
        if idx_fp_chs == -1: 
            if log: logger.warning('FORCE_PLATFORM:CHANNEL does not exist')
            return None        
        idx_analog_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        if idx_analog_labels == -1:
            if log: logger.warning('ANALOG:LABELS does not exist')
            return None       
        idx_analog_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
        if idx_analog_scale == -1:
            if log: logger.warning('ANALOG:SCALE does not exist')
            return None     
        idx_analog_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
        if idx_analog_offset == -1:
            if log: logger.warning('ANALOG:OFFSET does not exist')
            return None
        idx_analog_units = itf.GetParameterIndex('ANALOG', 'UNITS')
        if idx_analog_units == -1:
            if log: logger.warning('ANALOG:UNITS does not exist')
            n_analog_units = 0
        else:
            n_analog_units = itf.GetParameterLength(idx_analog_units)
        idx_point_units = itf.GetParameterIndex('POINT', 'UNITS')
        if idx_point_units == -1:
            if log: logger.warning('ANALOG:UNITS does not exist')
            point_unit = None
        else:
            point_unit = itf.GetParameterValue(idx_point_units, 0)
        point_scale = 1.0 if point_unit=='m' else 0.001
        analog_fps = get_analog_fps(itf, log=log)
        fp_params = get_fp_params(itf, log=log)
        fp_types = fp_params.get('TYPE', None)
        fp_origins = fp_params.get('ORIGIN', None)
        fp_corner_grps = fp_params.get('CORNERS', None)
        fp_chs = fp_params.get('CHANNEL', None)
        fp_cal_mats = fp_params.get('CAL_MATRIX', None)
        fp_output = {}
        for fp_idx in range(n_fp_used):
            fp_type = fp_types[fp_idx]            
            fp_org_raw = fp_origins[fp_idx]*point_scale
            fp_z_check = -1.0 if fp_org_raw[2]>0 else 1.0
            if fp_type == 1:
                o_x = 0.0
                o_y = 0.0
                o_z = (-1.0)*fp_org_raw[2]*fp_z_check
            elif fp_type in [2, 4]:
                o_x = (-1.0)*fp_org_raw[0]*fp_z_check
                o_y = (-1.0)*fp_org_raw[1]*fp_z_check
                o_z = (-1.0)*fp_org_raw[2]*fp_z_check
            elif fp_type == 3:
                o_x = 0.0
                o_y = 0.0
                o_z = (-1.0)*fp_org_raw[2]*fp_z_check
                fp_len_a = np.abs(fp_org_raw[0])
                fp_len_b = np.abs(fp_org_raw[1])
            fp_corners = fp_corner_grps[fp_idx]*point_scale
            fp_cen = np.mean(fp_corners, axis=0)
            fp_len_x = (np.linalg.norm(fp_corners[0]-fp_corners[1])+np.linalg.norm(fp_corners[3]-fp_corners[2]))*0.5
            fp_len_y = (np.linalg.norm(fp_corners[0]-fp_corners[3])+np.linalg.norm(fp_corners[1]-fp_corners[2]))*0.5
            fp_p0 = fp_cen
            fp_p1 = 0.5*(fp_corners[0]+fp_corners[3])
            fp_p2 = 0.5*(fp_corners[0]+fp_corners[1])
            fp_v0 = fp_p1-fp_p0
            fp_v1 = fp_p2-fp_p0
            fp_v0_u = fp_v0/np.linalg.norm(fp_v0)
            fp_v1_u = fp_v1/np.linalg.norm(fp_v1)
            fp_v2 = np.cross(fp_v0_u, fp_v1_u)
            fp_v2_u = fp_v2/np.linalg.norm(fp_v2)
            fp_v_z = fp_v2_u
            fp_v_x = fp_v0_u
            fp_v_y = np.cross(fp_v_z, fp_v_x)
            fp_rot_mat = np.column_stack([fp_v_x, fp_v_y, fp_v_z])
            chs = fp_chs[fp_idx]
            fp_data = {}
            ch_data = {}
            ch_unit_scale = {}
            gen_scale = get_analog_gen_scale(itf, log=log)
            sig_format = get_analog_format(itf, log=log)
            is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
            sig_offset_dtype = [np.int16, np.uint16][is_sig_unsigned]
            if filt_fc is None:
                filt_fcs = [None]*len(chs)
            elif type(filt_fc) in [int, float]:
                filt_fcs = [float(filt_fc)]*len(chs)
            elif type(filt_fc) in [list, tuple]:
                filt_fcs = list(filt_fc)
            elif type(filt_fc)==np.ndarray:
                if len(filt_fc)==1:
                    filt_fcs = [filt_fc.item()]*len(chs)
                else:
                    filt_fcs = filt_fc.tolist()
            for idx, ch in enumerate(chs):
                ch_idx = ch-1
                ch_name = itf.GetParameterValue(idx_analog_labels, ch_idx)
                ch_unit = ''
                if ch_idx < n_analog_units:
                    ch_unit = itf.GetParameterValue(idx_analog_units, ch_idx)
                ch_scale = np.float32(itf.GetParameterValue(idx_analog_scale, ch_idx))
                ch_offset = np.float32(sig_offset_dtype(itf.GetParameterValue(idx_analog_offset, ch_idx)))
                ch_val = (np.asarray(itf.GetAnalogDataEx(ch_idx, start_fr, end_fr, '0', 0, 0, '0'), dtype=np.float32)-ch_offset)*ch_scale*gen_scale
                # assign channel names
                if fp_type == 1:
                    # assume that the order of input analog channels are as follows:
                    # 'FX', 'FY', 'FZ', 'PX', 'PY', 'TZ'
                    ch_label = ['FX', 'FY', 'FZ', 'PX', 'PY', 'TZ'][idx]            
                elif fp_type in [2, 4]:
                    # assume that the order of input analog channels are as follows:
                    # 'FX', 'FY', 'FZ', 'MX', 'MY', 'MZ'
                    ch_label = ['FX', 'FY', 'FZ', 'MX', 'MY', 'MZ'][idx]  
                elif fp_type == 3:
                    # assume that the order of input analog channels are as follows:
                    # 'FX12', 'FX34', 'FY14', 'FY23', 'FZ1', 'FZ2', 'FZ3', 'FZ4'
                    ch_label = ['FX12', 'FX34', 'FY14', 'FY23', 'FZ1', 'FZ2', 'FZ3', 'FZ4'][idx]
                # assign channel scale factors
                if fp_type == 1:
                    if ch_label.startswith('F'):
                        # assume that the force unit is 'N'
                        ch_unit_scale[ch_label] = 1.0
                    elif ch_label.startswith('T'):
                        # assume that the torque unit is 'Nmm'
                        ch_unit_scale[ch_label] = 0.001
                        if ch_unit=='Nm': ch_unit_scale[ch_label] = 1.0
                    elif ch_label.startswith('P'):
                        # assume that the position unit is 'mm'
                        ch_unit_scale[ch_label] = 0.001
                        if ch_unit=='m': ch_unit_scale[ch_label] = 1.0
                elif fp_type in [2, 3, 4]:
                    if ch_label.startswith('F'):
                        # assume that the force unit is 'N'
                        ch_unit_scale[ch_label] = 1.0
                    elif ch_label.startswith('M'):
                        # assume taht the torque unit is 'Nmm'
                        ch_unit_scale[ch_label] = 0.001
                        if ch_unit=='Nm': ch_unit_scale[ch_label] = 1.0
                # assign channel values
                lp_fc = filt_fcs[idx]
                if lp_fc is None:
                    ch_data[ch_label] = ch_val
                else:
                    ch_data[ch_label] = np.float32(filt_bw_lp(ch_val, lp_fc, analog_fps, order=filt_order))
            if fp_type == 1:
                cop_l_x_in = ch_data['PX']*ch_unit_scale['PX']
                cop_l_y_in = ch_data['PY']*ch_unit_scale['PY']
                t_z_in = ch_data['TZ']*ch_unit_scale['TZ']
                fx = ch_data['FX']*ch_unit_scale['FX']
                fy = ch_data['FY']*ch_unit_scale['FY']
                fz = ch_data['FZ']*ch_unit_scale['FZ']
                mx = (cop_l_y_in-o_y)*fz+o_z*fy
                my = -o_z*fx-(cop_l_x_in-o_x)*fz
                mz = (cop_l_x_in-o_x)*fy-(cop_l_y_in-o_y)*fx+t_z_in
                f_raw = np.stack([fx, fy, fz], axis=1)
                m_raw = np.stack([mx, my, mz], axis=1)
            elif fp_type == 2:
                f_raw = np.stack([ch_data['FX']*ch_unit_scale['FX'], ch_data['FY']*ch_unit_scale['FY'], ch_data['FZ']*ch_unit_scale['FZ']], axis=1)
                m_raw = np.stack([ch_data['MX']*ch_unit_scale['MX'], ch_data['MY']*ch_unit_scale['MY'], ch_data['MZ']*ch_unit_scale['MZ']], axis=1)
            elif fp_type == 4:
                fp_cal_mat = fp_cal_mats[fp_idx]
                fm_local = np.stack([ch_data['FX'], ch_data['FY'], ch_data['FZ'], ch_data['MX'], ch_data['MY'], ch_data['MZ']], axis=1)
                fm_calib = np.dot(fp_cal_mat, fm_local.T).T
                f_raw = np.stack([fm_calib[:,0]*ch_unit_scale['FX'], fm_calib[:,1]*ch_unit_scale['FY'], fm_calib[:,2]*ch_unit_scale['FZ']], axis=1)
                m_raw = np.stack([fm_calib[:,3]*ch_unit_scale['MX'], fm_calib[:,4]*ch_unit_scale['MY'], fm_calib[:,5]*ch_unit_scale['MZ']], axis=1)
            elif fp_type == 3:
                fx12 = ch_data['FX12']*ch_unit_scale['FX12']
                fx34 = ch_data['FX34']*ch_unit_scale['FX34']
                fy14 = ch_data['FY14']*ch_unit_scale['FY14']
                fy23 = ch_data['FY23']*ch_unit_scale['FY23']
                fz1 = ch_data['FZ1']*ch_unit_scale['FZ1']
                fz2 = ch_data['FZ2']*ch_unit_scale['FZ2']
                fz3 = ch_data['FZ3']*ch_unit_scale['FZ3']
                fz4 = ch_data['FZ4']*ch_unit_scale['FZ4']
                fx = fx12+fx34
                fy = fy14+fy23
                fz = fz1+fz2+fz3+fz4
                mx = fp_len_b*(fz1+fz2-fz3-fz4)
                my = fp_len_a*(-fz1+fz2+fz3-fz4)
                mz = fp_len_b*(-fx12+fx34)+fp_len_a*(fy14-fy23)
                f_raw = np.stack([fx, fy, fz], axis=1)
                m_raw = np.stack([mx, my, mz], axis=1)
            zero_vals = np.zeros((f_raw.shape[0]), dtype=np.float32)
            fm_skip_mask = np.abs(f_raw[:,2])<=threshold
            f_sensor_local = f_raw.copy()
            m_sensor_local = m_raw.copy()
            # filter local values by threshold
            f_sensor_local[fm_skip_mask,:] = 0.0
            m_sensor_local[fm_skip_mask,:] = 0.0
            f_x = f_sensor_local[:,0]
            f_y = f_sensor_local[:,1]
            f_z = f_sensor_local[:,2]
            m_x = m_sensor_local[:,0]
            m_y = m_sensor_local[:,1]
            m_z = m_sensor_local[:,2]
            with np.errstate(invalid='ignore'):
                f_z_adj = np.where(fm_skip_mask, np.inf, f_z)
                cop_l_x = np.where(fm_skip_mask, np.nan, np.clip((-m_y+(-o_z)*f_x)/f_z_adj+o_x, -fp_len_x*0.5, fp_len_x*0.5))
                cop_l_y = np.where(fm_skip_mask, np.nan, np.clip((m_x+(-o_z)*f_y)/f_z_adj+o_y, -fp_len_y*0.5, fp_len_y*0.5))
                cop_l_z = np.where(fm_skip_mask, np.nan, zero_vals)
                if cop_nan_to_num:
                    cop_l_x = np.nan_to_num(cop_l_x)
                    cop_l_y = np.nan_to_num(cop_l_y)
                    cop_l_z = np.nan_to_num(cop_l_z)
            t_z = m_z-(cop_l_x-o_x)*f_y+(cop_l_y-o_y)*f_x
            # values for the force plate local output
            m_cop_local = np.stack([zero_vals, zero_vals, t_z], axis=1)
            cop_surf_local = np.stack([cop_l_x, cop_l_y, cop_l_z], axis=1)
            f_surf_local = f_sensor_local
            m_surf_local = np.cross(np.array([o_x, o_y, o_z], dtype=np.float32), f_sensor_local)+m_sensor_local
            # values for the force plate global output
            m_cop_global = np.dot(fp_rot_mat, m_cop_local.T).T
            cop_surf_global = np.dot(fp_rot_mat, cop_surf_local.T).T
            f_surf_global = np.dot(fp_rot_mat, f_surf_local.T).T
            m_surf_global = np.dot(fp_rot_mat, m_surf_local.T).T
            # values for the lab output
            m_cop_lab = m_cop_global
            cop_lab = fp_cen+cop_surf_global
            f_cop_lab = f_surf_global
            # prepare return values        
            fp_data.update({'F_SURF_LOCAL': f_surf_local})
            fp_data.update({'M_SURF_LOCAL': m_surf_local})
            fp_data.update({'COP_SURF_LOCAL': cop_surf_local})
            fp_data.update({'F_SURF_GLOBAL': f_surf_global})
            fp_data.update({'M_SURF_GLOBAL': m_surf_global})
            fp_data.update({'COP_SURF_GLOBAL': cop_surf_global})
            fp_data.update({'F_COP_LAB': f_cop_lab})            
            fp_data.update({'M_COP_LAB': m_cop_lab})
            fp_data.update({'COP_LAB': cop_lab})
            if fp_type == 1:
                fp_data.update({'COP_LOCAL_INPUT': np.stack([cop_l_x_in, cop_l_y_in, zero_vals], axis=1)})
            fp_output.update({fp_idx: fp_data})   
        return fp_output
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    
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
    try:
        mkr_idx = get_marker_index(itf, mkr_name_old, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            err_msg = f'Unable to get the index of "{mkr_name_old}"'
            raise ValueError(err_msg)
        par_idx = itf.GetParameterIndex('POINT', 'LABELS')
        if par_idx == -1:
            err_msg = 'POINT:LABELS does not exist'
            raise RuntimeError(err_msg)
        ret = itf.SetParameterValue(par_idx, mkr_idx, mkr_name_new)
        if log:
            logger.info(f'Changing the marker name from "{mkr_name_old}" to "{mkr_name_new}": {["FAILURE", "SUCCESS"][ret]}')
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise        

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
    try:
        sig_idx = get_analog_index(itf, sig_name_old, log=log)
        if sig_idx == -1 or sig_idx is None:
            err_msg = f'Unable to get the index of "{sig_name_old}"'
            raise ValueError(err_msg)
        par_idx = itf.GetParameterIndex('ANALOG', 'LABELS')
        if par_idx == -1:
            err_msg = 'ANALOG:LABELS does not exist'
            raise RuntimeError(err_msg)
        ret = itf.SetParameterValue(par_idx, sig_idx, sig_name_new)
        if log:
            logger.info(f'Changing the signal name from "{sig_name_old}" to "{sig_name_new}": {["FAILURE", "SUCCESS"][ret]}')    
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise

def resize_char_type_param(itf, grp_name, param_name, new_str_len, log=False):
    """
    Resize a char type parameter's first dimension.
    
    This function only works with 2 dimensional char type parameters.
    
    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    grp_name : str
        Group name.
    param_name : str
        Parameter name.
    new_str_len : int
        Size(first dimension) of a new string for parameter.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        idx_par = itf.GetParameterIndex(grp_name, param_name)
        if idx_par == -1:
            err_msg = f'{grp_name}:{param_name} does not exist'
            raise ValueError(err_msg)
        par_type = itf.GetParameterType(idx_par)
        if par_type != -1:
            err_msg = '{grp_name}:{param_name} is not a char type parameter'
            raise ValueError(err_msg)
        par_num_dim = itf.GetParameterNumberDim(idx_par)
        if par_num_dim != 2:
            err_msg = f'{grp_name}:{param_name} is not a 2 dimensional parameter'
            raise ValueError(err_msg)
        par_dim_old = [itf.GetParameterDimension(idx_par, j) for j in range(par_num_dim)]
        par_desc = itf.GetParameterDescription(idx_par)
        n_par = itf.GetParameterLength(idx_par)
        par_data = []
        for i in range(n_par):
            par_data.append(itf.GetParameterValue(idx_par, i))                
        max_str_len = max([len(x) for x in par_data])
        if new_str_len < max_str_len:
            err_msg = f'"new_str_len" should be equal or greater than {max_str_len}'
            raise ValueError(err_msg)            
        par_dim = par_dim_old
        par_dim[0] = new_str_len
        # Use a pure python list of strings for pythoncom.VT_ARRAY|pythoncom.VT_BSTR instead of a ndarray
        var_par_dim = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2, par_dim)
        var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_BSTR, par_data)
        # Delete existing parameter
        ret = [False, True][itf.DeleteParameter(idx_par)]
        if not ret:
            err_msg = f'Failed to delete existing {grp_name}:{param_name}'
            raise RuntimeError(err_msg)             
        # Create new parameter
        idx_par = itf.AddParameter(param_name, par_desc, grp_name, np.uint8(0), par_type, par_num_dim, var_par_dim, var_par_data)
        if idx_par == -1:
            err_msg = f'Failed to create new {grp_name}:{param_name}'
            raise RuntimeError(err_msg)
        else:
            return True    
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise
        
def adjust_param_items(itf, grp_name, param_name, recreate_param=False, keep_str_len=True, log=False):
    """
    Adjust a specific group parameter's length to be compatible with USED parameter.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    grp_name : str
        Group name.
    param_name : str
        Parameter name.
    recreate_param : bool, optional
        Whether to recreate the parameter or not. The default is False.
    keep_str_len : bool, optional
        Whether to keep the first dimension (string length) of the parameter if you recreate a parameter. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        if grp_name not in ['POINT', 'ANALOG']:
            err_msg = f'This function only works with either POINT or ANALOG group'
            raise ValueError(err_msg)
        if grp_name=='POINT' and param_name not in ['DESCRIPTIONS', 'LABELS']:
            err_msg = f'{grp_name}:{param_name} is not supported with this function'
            raise ValueError(err_msg)
        if grp_name=='ANALOG' and param_name not in ['DESCRIPTIONS', 'LABELS', 'UNITS', 'SCALE', 'OFFSET']:
            err_msg = f'{grp_name}:{param_name} is not supported with this function'
            raise ValueError(err_msg)            
        if param_name == 'DESCRIPTIONS':
            idx_used = itf.GetParameterIndex(grp_name, 'USED')
            n_used = itf.GetParameterValue(idx_used, 0)
            idx_par = itf.GetParameterIndex(grp_name, param_name)
            n_par = itf.GetParameterLength(idx_par)
            desc_len_max = itf.GetParameterDimension(idx_par, 0)
            if n_par == n_used:
                if log: logger.debug(f'{grp_name}:{param_name} has as same number of items as {grp_name}:USED')
                return False            
            elif n_par < n_used:
                if log: logger.debug(f'{grp_name}:{param_name} has fewer items than {grp_name}:USED')
                # for i in range(n_par, n_used):
                #     ret = itf.AddParameterData(idx_par, 1)
                #     par_len = itf.GetParameterLength(idx_par)
                #     var_desc = win32.VARIANT(pythoncom.VT_BSTR, str(par_len))
                #     ret = itf.SetParameterValue(idx_par, par_len-1, var_desc)
                ret = itf.AddParameterData(idx_par, n_used-n_par)
                for i in range(n_par, n_used):
                    str_desc = str(i+1) if len(str(i+1)) <= desc_len_max else ''
                    var_desc = win32.VARIANT(pythoncom.VT_BSTR, str_desc)
                    ret = itf.SetParameterValue(idx_par, i, var_desc)                    
            else:
                if log: logger.debug(f'{grp_name}:{param_name} has more items than {grp_name}:USED, so all unused items will be deleted')
                if recreate_param:
                    par_desc = itf.GetParameterDescription(idx_par)
                    par_num_dim = itf.GetParameterNumberDim(idx_par)
                    par_dim_old = [itf.GetParameterDimension(idx_par, j) for j in range(par_num_dim)]
                    par_data = []
                    for i in range(n_used):
                        par_data.append(itf.GetParameterValue(idx_par, i))
                    par_type = itf.GetParameterType(idx_par)
                    size_par_ideal = math.ceil(float(par_dim_old[0]*par_dim_old[1])/float(n_used))
                    size_par = [size_par_ideal, par_dim_old[0]][keep_str_len]
                    par_dim = [size_par, n_used]
                    var_par_dim = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2, par_dim)
                    # Use a pure python list of strings for pythoncom.VT_ARRAY|pythoncom.VT_BSTR instead of a ndarray
                    var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_BSTR, par_data)
                    # Delete existing parameter
                    ret = [False, True][itf.DeleteParameter(idx_par)]
                    if not ret:
                        err_msg = f'Failed to delete existing {grp_name}:{param_name}'
                        raise RuntimeError(err_msg)
                    # Create new parameter
                    idx_par = itf.AddParameter(param_name, par_desc, grp_name, np.uint8(0), par_type, par_num_dim, var_par_dim, var_par_data)
                    if idx_par == -1:
                        err_msg = f'Failed to create new {grp_name}:{param_name}'
                        raise RuntimeError(err_msg)
                    else:
                        return True
                else:
                    for i in range(n_par-1, n_used-1, -1):
                        ret = itf.RemoveParameterData(idx_par, i)
                        if ret == 0:
                            err_msg = f'Failed to delete {i+1}th item under {grp_name}:{param_name}'
                            raise RuntimeError(err_msg)
                    return True
        else:
            idx_used = itf.GetParameterIndex(grp_name, 'USED')
            n_used = itf.GetParameterValue(idx_used, 0)
            idx_par = itf.GetParameterIndex(grp_name, param_name)
            n_par = itf.GetParameterLength(idx_par)
            if n_par == n_used:
                if log: logger.debug(f'{grp_name}:{param_name} has as same number of items as {grp_name}:USED')
                return False            
            elif n_par < n_used:
                err_msg = f'Number of item under {grp_name}:{param_name} is less than {grp_name}:USED'
                raise RuntimeError(err_msg)
            else:
                if log: logger.debug(f'{grp_name}:{param_name} has more items than {grp_name}:USED, so all unused items will be deleted')
                if recreate_param:
                    par_desc = itf.GetParameterDescription(idx_par)
                    par_num_dim = itf.GetParameterNumberDim(idx_par)
                    par_dim_old = [itf.GetParameterDimension(idx_par, j) for j in range(par_num_dim)]
                    par_data = []
                    for i in range(n_used):
                        par_data.append(itf.GetParameterValue(idx_par, i))                
                    par_type = itf.GetParameterType(idx_par)
                    if par_type == -1:
                        size_par_ideal = math.ceil(float(par_dim_old[0]*par_dim_old[1])/float(n_used))
                        size_par = [size_par_ideal, par_dim_old[0]][keep_str_len]
                        par_dim = [size_par, n_used]
                    else:
                        par_dim = [n_used]
                    var_par_dim = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2, par_dim)
                    # Use a pure python list of strings for pythoncom.VT_ARRAY|pythoncom.VT_BSTR instead of a ndarray
                    if par_type == -1:
                        var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_BSTR, par_data)
                    elif par_type == 1:
                        var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I1, par_data)
                    elif par_type == 2:
                        var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2, par_data)
                    elif par_type == 4:
                        var_par_data = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, par_data)
                    else:
                        err_msg = f'Unknown data type from {grp_name}:{param_name}'
                        raise RuntimeError(err_msg)              
                    # Delete existing parameter
                    ret = [False, True][itf.DeleteParameter(idx_par)]
                    if not ret:
                        err_msg = f'Failed to delete existing {grp_name}:{param_name}'
                        raise RuntimeError(err_msg)             
                    # Create new parameter
                    idx_par = itf.AddParameter(param_name, par_desc, grp_name, np.uint8(0), par_type, par_num_dim, var_par_dim, var_par_data)
                    if idx_par == -1:
                        err_msg = f'Failed to create new {grp_name}:{param_name}'
                        raise RuntimeError(err_msg)
                    else:
                        return True                    
                else:
                    for i in range(n_par-1, n_used-1, -1):
                        ret = itf.RemoveParameterData(idx_par, i)
                        if ret == 0:
                            err_msg = f'Failed to delete {i+1}th item under {grp_name}:{param_name}'
                            raise RuntimeError(err_msg)                       
                    return True
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise          
    
def auto_adjust_params(itf, recreate_param=False, keep_str_len=True, log=False):
    """
    Adjust several group parameters' items automatically, especially for POINT and ANALOG groups.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    recreate_param : bool, optional
        Whether to recreate the parameter or not. The default is False.
    keep_str_len : bool, optional
        Whether to keep the first dimension (string length) of the parameter if you recreate a parameter. The default is True.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    None.

    """
    try:
        adjust_param_items(itf, 'POINT', 'LABELS', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'POINT', 'DESCRIPTIONS', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'ANALOG', 'SCALE', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'ANALOG', 'OFFSET', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'ANALOG', 'UNITS', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'ANALOG', 'LABELS', recreate_param, keep_str_len, log=log)
        adjust_param_items(itf, 'ANALOG', 'DESCRIPTIONS', recreate_param, keep_str_len, log=log)
        return None
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise

def add_group(itf, grp_name, grp_desc=None, grp_lock=False, log=False):
    """
    Add a group.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    grp_name : str
        Group name.
    grp_desc : str, optional
        Group description. The default is None.
    grp_lock : bool, optional
        Whether to lock the group or not. The default is False.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    ret : int
        The index of newly created group.

    """
    try:
        desc = '' if grp_desc is None else grp_desc
        lock = ['0', '1'][grp_lock]
        ret = itf.AddGroup(0, grp_name, desc, lock)
        if ret == -1:
            err_msg = 'Group could not be added'
            raise RuntimeError(err_msg)
        else:
            return ret            
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise        

def add_param(itf, grp_name, param_name, param_data, param_desc=None, make_new_grp=False, log=False):
    """
    Add a parameter under the specified group.
    
    This function only supports string, integer and float data types.
    Parameter data should be in form of either single value or 1-d array like strctures (list, tuple, numpy array).

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    grp_name : str
        Group name.
    param_name : str
        Parameter name.
    param_data : str, int, float, list, tuple, ndarray
        Parameter data. This argument should be either single value or 1 dimensional list, tuple, numpy array.
    param_desc : str, optional
        Parameter description. The default is None.
    make_new_grp : bool, optional
        Whether to make a new group if the given group name does not exist. The default is False.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    ret : int
        The index of newly created parameter.

    """
    try:
        idx_grp = itf.GetGroupIndex(grp_name)
        if idx_grp == -1:
            if make_new_grp:
                # itf.AddGroup(0, grp_name, '', '0')
                add_group(itf, grp_name, grp_lock=False, log=log)
            else:
                err_msg = f'"{grp_name}" group does not exist'
                raise RuntimeError(err_msg)
        idx_par = itf.GetParameterIndex(grp_name, param_name)
        if idx_par != -1:
            err_msg = f'"{grp_name}:{param_name}" parameter already exists'
            raise RuntimeError(err_msg)
        par_data_type = type(param_data)
        if par_data_type == str:
            par_dtype = -1
            par_num_dim = 2
            par_dim = [len(param_data), 1]
            par_data = [param_data]
            var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_BSTR
        elif par_data_type == int:
            par_dtype = 2
            par_num_dim = 0
            par_dim = []
            par_data = [param_data]
            var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_I2
        elif par_data_type == float:
            par_dtype = 4
            par_num_dim = 0
            par_dim = []
            par_data = [param_data]
            var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_R4
        elif par_data_type==list or par_data_type==tuple or par_data_type==np.ndarray:
            if par_data_type==list or par_data_type==tuple:
                par_data = list(param_data)
            elif par_data_type == np.ndarray:
                par_data = param_data.tolist()
            if all(isinstance(x, str) for x in par_data):
                par_dtype = -1
                par_num_dim = 2
                par_dim = []
                par_dim.append(max([len(x) for x in par_data]))
                par_dim.append(len(par_data))
                var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_BSTR
            elif all(isinstance(x, int) for x in par_data):
                par_dtype = 2
                par_num_dim = 1
                par_dim = []
                par_dim.append(len(par_data))
                var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_I2
            elif all(isinstance(x, float) for x in par_data):
                par_dtype = 4
                par_num_dim = 1
                par_dim = []
                par_dim.append(len(par_data))
                var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_R4
            elif any(isinstance(x, float) for x in par_data) and all(~isinstance(x, str) for x in par_data):
                par_dtype = 4
                par_num_dim = 1
                par_dim = []
                par_dim.append(len(par_data))
                var_par_dtype = pythoncom.VT_ARRAY|pythoncom.VT_R4                
            else:
                err_msg = 'Unsupported data type'
                raise ValueError(err_msg)
        else:
            err_msg = 'Unsupported data type'
            raise ValueError(err_msg)
        var_par_dim = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2, par_dim)
        var_par_data = win32.VARIANT(var_par_dtype, par_data)
        par_desc = '' if param_desc is None else param_desc
        ret = itf.AddParameter(param_name, par_desc, grp_name, np.uint8(0), par_dtype, par_num_dim, var_par_dim, var_par_data)
        if ret == -1:
            err_msg = 'Parameter could not be added'
            raise RuntimeError(err_msg)
        else:
            return ret
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise        
    
def add_marker(itf, mkr_name, mkr_coords, mkr_resid=None, mkr_desc=None, adjust_params=False, log=False):
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
        A numpy array of new marker coordinates. This is assumed as a scaled one.
    mkr_resid : numpy array or None, optional
        A numpy array of new marker residuals. The default is None.
    mkr_desc : str or None, optional
        Description of a new marker.
    adjust_params : bool, optional
        Whether to adjust the lengths of other related parameters. The default is False.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True of False.

    """
    try:
        if log: logger.debug(f'Start adding a new "{mkr_name}" marker ...')
        start_fr = get_first_frame(itf, log=log)
        n_frs = get_num_frames(itf, log=log)
        if not (mkr_coords.ndim==2 and mkr_coords.shape[0]==n_frs and mkr_coords.shape[1]==3):
            err_msg = 'Not valid dimensions of input marker coordinates'
            raise ValueError(err_msg)
        if mkr_resid is not None:
            if not (mkr_resid.ndim==1 and mkr_resid.shape[0]==n_frs):
                err_msg = 'Not valid dimensions of input marker residuals'
                raise ValueError(err_msg)
        # Adjust POINT group parameters
        if adjust_params:
            adjust_param_items(itf, 'POINT', 'LABELS', recreate_param=False, keep_str_len=True, log=log)
            adjust_param_items(itf, 'POINT', 'DESCRIPTIONS', recreate_param=False, keep_str_len=True, log=log)
        ret = 0
        # Check 'POINT:USED'
        idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
        n_pt_used_before = itf.GetParameterValue(idx_pt_used, 0)
        # Check 'POINT:LABELS'
        idx_pt_labels = itf.GetParameterIndex('POINT', 'LABELS')
        n_pt_labels_before = itf.GetParameterLength(idx_pt_labels)
        # Skip if 'POINT:USED' and 'POINT:LABELS' have different numbers
        if n_pt_used_before != n_pt_labels_before:
            err_msg0 = 'This function only works if POINT:USED is as same as the number of items under POINT:LABELS'
            err_msg1 = ', so please try with "adjust_params" as "True"'
            err_msg = err_msg0+err_msg1
            raise RuntimeError(err_msg)
        # Add an parameter to 'POINT:LABELS'
        ret = itf.AddParameterData(idx_pt_labels, 1)
        if ret == 0:
            err_msg = f'Failed to add an item under POINT:LABELS'
            raise RuntimeError(err_msg)     
        n_pt_labels = itf.GetParameterLength(idx_pt_labels)
        var_mkr_name = win32.VARIANT(pythoncom.VT_BSTR, mkr_name)
        ret = itf.SetParameterValue(idx_pt_labels, n_pt_labels-1, var_mkr_name)
        if ret == 0:
            err_msg = f'Failed to set the value of an item under POINT:LABELS'
            raise RuntimeError(err_msg)
        # Add a null parameter in the 'POINT:DESCRIPTIONS' section
        idx_pt_desc = itf.GetParameterIndex('POINT', 'DESCRIPTIONS')
        ret = itf.AddParameterData(idx_pt_desc, 1)
        if ret == 0:
            err_msg = f'Failed to add an item under POINT:DESCRIPTIONS'
            raise RuntimeError(err_msg)
        n_pt_desc = itf.GetParameterLength(idx_pt_desc)
        mkr_desc_adjusted = '' if mkr_desc is None else mkr_desc
        var_mkr_desc = win32.VARIANT(pythoncom.VT_BSTR, mkr_desc_adjusted)
        ret = itf.SetParameterValue(idx_pt_desc, n_pt_desc-1, var_mkr_desc)
        if ret == 0:
            err_msg = f'Failed to set the value of an item under POINT:DESCRIPTIONS'
            raise RuntimeError(err_msg)        
        # Add a marker
        new_mkr_idx = itf.AddMarker()
        n_mkrs = itf.GetNumber3DPoints()
        mkr_null_masks = np.any(np.isnan(mkr_coords), axis=1)
        mkr_resid_adjusted = np.zeros((n_frs, ), dtype=np.float32) if mkr_resid is None else np.array(mkr_resid, dtype=np.float32)
        mkr_resid_adjusted[mkr_null_masks] = -1
        # mkr_masks = np.array(['0000000']*n_frs, dtype = np.string_)
        mkr_masks = ['0000000']*n_frs
        mkr_scale = get_marker_scale(itf, log=log)
        if mkr_scale is None:
            err_msg = f'Unable to get the marker scale factor'
            raise RuntimeError(err_msg)        
        is_c3d_float = mkr_scale < 0
        is_c3d_float2 = [False, True][itf.GetDataType()-1]
        if is_c3d_float != is_c3d_float2:
            if log: logger.debug('C3D data type is determined by POINT:SCALE')
        mkr_dtype = [np.int16, np.float32][is_c3d_float]    
        scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
        if is_c3d_float:
            mkr_coords_unscaled = np.asarray(np.nan_to_num(mkr_coords), dtype=mkr_dtype)
        else:
            mkr_coords_unscaled = np.asarray(np.round(np.nan_to_num(mkr_coords)/scale_size), dtype=mkr_dtype)
        dtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_c3d_float]
        dtype_arr = pythoncom.VT_ARRAY|dtype
        for i in range(3):
            # var_pos = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i])
            var_pos = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i].tolist())
            ret = itf.SetPointDataEx(n_mkrs-1, i, start_fr, var_pos)
            if ret == 0:
                err_msg = f'Failed to set the data for a new marker'
                raise RuntimeError(err_msg)
        # var_resid = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, mkr_resid_adjusted)
        var_resid = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, mkr_resid_adjusted.tolist())
        ret = itf.SetPointDataEx(n_mkrs-1, 3, start_fr, var_resid)
        if ret == 0:
            err_msg = f'Failed to set the data for a new marker'
            raise RuntimeError(err_msg)        
        var_masks = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_BSTR, mkr_masks)
        ret = itf.SetPointDataEx(n_mkrs-1, 4, start_fr, var_masks)
        if ret == 0:
            err_msg = f'Failed to set the data for a new marker'
            raise RuntimeError(err_msg)        
        var_const = win32.VARIANT(dtype, 1)
        for i in range(3):
            for idx, val in enumerate(mkr_coords_unscaled[:,i]):
                if val == 1:
                    ret = itf.SetPointData(n_mkrs-1, i, start_fr+idx, var_const)
                    if ret == 0:
                        err_msg = f'Failed to set the data for a new marker'
                        raise RuntimeError(err_msg)
        # Increase 'POINT:USED' by 1
        idx_pt_used = itf.GetParameterIndex('POINT', 'USED')
        n_pt_used_after = itf.GetParameterValue(idx_pt_used, 0)
        if n_pt_used_after != (n_pt_used_before+1):
            if log: log.debug('POINT:USED was not properly updated so that manual update will be executed')
            ret = itf.SetParameterValue(idx_pt_used, 0, (n_pt_used_before+1))
            if ret == 0:
                err_msg = f'Failed to set the value of POINT:USED'
                raise RuntimeError(err_msg)            
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise        

def add_analog(itf, sig_name, sig_value, sig_unit, sig_scale=1.0, sig_offset=0, sig_gain=0, sig_desc=None, adjust_params=False, log=False):
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
        A new analog channel value. This is assumed as a scaled one.
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
    adjust_params : bool, optional
        Whether to adjust the lengths of other related parameters. The default is False.        
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        if log: logger.debug(f'Start adding a new "{sig_name}" analog channel ...')
        start_fr = get_first_frame(itf, log=log)
        n_frs = get_num_frames(itf, log=log)
        av_ratio = get_analog_video_ratio(itf, log=log)
        if sig_value.ndim!=1 or sig_value.shape[0]!=(n_frs*av_ratio):
            err_msg = 'The dimension of the input signal value is not compatible'
            raise ValueError(err_msg)          
        # Adjust ANALOG group parameters
        if adjust_params:
            adjust_param_items(itf, 'ANALOG', 'SCALE', recreate_param=False, keep_str_len=True, log=log)
            adjust_param_items(itf, 'ANALOG', 'OFFSET', recreate_param=False, keep_str_len=True, log=log)
            adjust_param_items(itf, 'ANALOG', 'UNITS', recreate_param=False, keep_str_len=True, log=log)
            adjust_param_items(itf, 'ANALOG', 'LABELS', recreate_param=False, keep_str_len=True, log=log)
            adjust_param_items(itf, 'ANALOG', 'DESCRIPTIONS', recreate_param=False, keep_str_len=True, log=log)      
        # Check 'ANALOG:USED'
        idx_an_used = itf.GetParameterIndex('ANALOG', 'USED')
        n_an_used_before = itf.GetParameterValue(idx_an_used, 0) 
        # Check 'ANALOG:LABELS'
        idx_an_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        n_an_labels_before = itf.GetParameterLength(idx_an_labels)
        # Skip if 'ANALOG:USED' and 'ANALOG:LABELS' have different numbers
        if n_an_used_before != n_an_labels_before:
            err_msg0 = 'This function only works if ANALOG:USED is as same as the number of items under ANALOG:LABELS'
            err_msg1 = ', so please try with "adjust_params" as "True"'
            err_msg = err_msg0+err_msg1
            raise RuntimeError(err_msg)
        # Add an parameter to the 'ANALOG:LABELS' section
        idx_an_labels = itf.GetParameterIndex('ANALOG', 'LABELS')
        ret = itf.AddParameterData(idx_an_labels, 1)
        n_an_labels = itf.GetParameterLength(idx_an_labels)
        ret = itf.SetParameterValue(idx_an_labels, n_an_labels-1, win32.VARIANT(pythoncom.VT_BSTR, sig_name))
        # Add an parameter to the 'ANALOG:UNITS' section
        idx_an_units = itf.GetParameterIndex('ANALOG', 'UNITS')
        ret = itf.AddParameterData(idx_an_units, 1)
        n_an_units = itf.GetParameterLength(idx_an_units)
        ret = itf.SetParameterValue(idx_an_units, n_an_units-1, win32.VARIANT(pythoncom.VT_BSTR, sig_unit))      
        # Add an parameter to the 'ANALOG:SCALE' section
        idx_an_scale = itf.GetParameterIndex('ANALOG', 'SCALE')
        ret = itf.AddParameterData(idx_an_scale, 1)
        n_an_scale = itf.GetParameterLength(idx_an_scale)
        ret = itf.SetParameterValue(idx_an_scale, n_an_scale-1, win32.VARIANT(pythoncom.VT_R4, sig_scale))
        # Add an parameter to the 'ANALOG:OFFSET' section
        idx_an_offset = itf.GetParameterIndex('ANALOG', 'OFFSET')
        ret = itf.AddParameterData(idx_an_offset, 1)
        n_an_offset = itf.GetParameterLength(idx_an_offset)
        sig_format = get_analog_format(itf, log=log)
        is_sig_unsigned = (sig_format is not None) and (sig_format.upper()=='UNSIGNED')
        sig_offset_comtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_sig_unsigned]
        sig_offset_dtype = [np.int16, np.uint16][is_sig_unsigned]
        ret = itf.SetParameterValue(idx_an_offset, n_an_offset-1, win32.VARIANT(sig_offset_comtype, sig_offset))
        # Check for 'ANALOG:GAIN' section and add 0 if it exists
        idx_an_gain = itf.GetParameterIndex('ANALOG', 'GAIN')
        if idx_an_gain != -1:
            ret = itf.AddParameterData(idx_an_gain, 1)
            n_an_gain = itf.GetParameterLength(idx_an_gain)
            ret = itf.SetParameterValue(idx_an_gain, n_an_gain-1, win32.VARIANT(pythoncom.VT_I2, sig_gain))    
        # Add an parameter to the 'ANALOG:DESCRIPTIONS' section
        idx_an_desc = itf.GetParameterIndex('ANALOG', 'DESCRIPTIONS')
        ret = itf.AddParameterData(idx_an_desc, 1)
        n_an_desc = itf.GetParameterLength(idx_an_desc)
        sig_desc_in = sig_name if sig_desc is None else sig_desc
        ret = itf.SetParameterValue(idx_an_desc, n_an_desc-1, win32.VARIANT(pythoncom.VT_BSTR, sig_desc_in))
        # Create an analog channel
        idx_new_an_ch = itf.AddAnalogChannel()
        # n_an_chs = itf.GetAnalogChannels()
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_value_unscaled = np.asarray(sig_value, dtype=np.float32)/(np.float32(sig_scale)*gen_scale)+np.float32(sig_offset_dtype(sig_offset))
        # variant = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, sig_value_unscaled)
        variant = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, sig_value_unscaled.tolist())
        ret = itf.SetAnalogDataEx(idx_new_an_ch, start_fr, variant)
        # Increase the value 'ANALOG:USED' by 1
        idx_an_used = itf.GetParameterIndex('ANALOG', 'USED')
        n_an_used_after = itf.GetParameterValue(idx_an_used, 0)
        if n_an_used_after != (n_an_used_before+1):
            if log: log.debug('ANALOG:USED was not properly updated so that manual update will be executed')
            ret = itf.SetParameterValue(idx_an_used, 0, (n_an_used_before+1))
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise

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
    try:
        if start_frame < get_first_frame(itf, log=log):
            err_msg = f'"start_frame" number should be equal or greater than {get_first_frame(itf)}'
            raise ValueError(err_msg)
        elif start_frame >= get_last_frame(itf, log=log):
            err_msg = f'"start_frame" number should be less than {get_last_frame(itf)}'
            raise ValueError(err_msg)
        n_frs_updated = itf.DeleteFrames(start_frame, num_frames)
        return n_frs_updated
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise        

def set_marker_pos(itf, mkr_name, mkr_coords, start_frame=None, end_frame=None, log=False):
    """
    Set the coordinates of a marker.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    mkr_coords : numpy array
        Marker coordinates.
    start_frame: None or int, optional
        User-defined start frame. The default is None.
    end_frame: None or int, optional
        User-defined end frame. The default is None.        
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            err_msg = 'Given "start_frame" is not proper'
            raise ValueError(err_msg)
        n_frs = end_fr-start_fr+1
        if mkr_coords.ndim != 2 or mkr_coords.shape[0] != n_frs:
            err_msg = 'The dimension of the input is not compatible'
            raise ValueError(err_msg)    
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            err_msg = f'Unable to get the index of "{mkr_name}"'
            raise ValueError(err_msg)
        mkr_scale = get_marker_scale(itf, log=log)
        if mkr_scale is None:
            err_msg = f'Unable to get the marker scale factor'
            raise RuntimeError(err_msg)
        is_c3d_float = mkr_scale < 0
        is_c3d_float2 = [False, True][itf.GetDataType()-1]
        if is_c3d_float != is_c3d_float2:
            if log: logger.debug('C3D data type is determined by POINT:SCALE')
        mkr_dtype = [np.int16, np.float32][is_c3d_float]
        scale_size = [np.fabs(mkr_scale), np.float32(1.0)][is_c3d_float]
        if is_c3d_float:
            mkr_coords_unscaled = np.asarray(np.nan_to_num(mkr_coords), dtype=mkr_dtype)
        else:
            mkr_coords_unscaled = np.asarray(np.round(np.nan_to_num(mkr_coords)/scale_size), dtype=mkr_dtype)
        dtype = [pythoncom.VT_I2, pythoncom.VT_R4][is_c3d_float]
        dtype_arr = pythoncom.VT_ARRAY|dtype
        for i in range(3):
            # variant = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i])
            variant = win32.VARIANT(dtype_arr, mkr_coords_unscaled[:,i].tolist())
            ret = itf.SetPointDataEx(mkr_idx, i, start_fr, variant)
        var_const = win32.VARIANT(dtype, 1)
        for i in range(3):
            for idx, val in enumerate(mkr_coords_unscaled[:,i]):
                if val == 1:
                    ret = itf.SetPointData(mkr_idx, i, start_fr+idx, var_const)
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    except RuntimeError as err:
        if log: logger.error(err)
        raise        
    
def set_marker_resid(itf, mkr_name, mkr_resid, start_frame=None, end_frame=None, log=False):
    """
    Set the residual of a marker.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    mkr_name : str
        Marker name.
    mkr_resid : numpy array
        Marker residuals.
    start_frame: None or int, optional
        User-defined start frame. The default is None.
    end_frame: None or int, optional
        User-defined end frame. The default is None.            
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            err_msg = 'Given "start_frame" is not proper'
            raise ValueError(err_msg)        
        n_frs = end_fr-start_fr+1
        if mkr_resid.ndim != 1 or mkr_resid.shape[0] != n_frs:
            err_msg = 'The dimension of the input is not compatible'
            raise ValueError(err_msg)        
        mkr_idx = get_marker_index(itf, mkr_name, log=log)
        if mkr_idx == -1 or mkr_idx is None:
            err_msg = f'Unable to get the index of "{mkr_name}"'
            raise ValueError(err_msg)        
        dtype = pythoncom.VT_R4
        dtype_arr = pythoncom.VT_ARRAY|dtype
        # variant = win32.VARIANT(dtype_arr, mkr_resid)
        variant = win32.VARIANT(dtype_arr, mkr_resid.tolist())
        ret = itf.SetPointDataEx(mkr_idx, 3, start_fr, variant)
        var_const = win32.VARIANT(dtype, 1)
        for idx, val in enumerate(mkr_resid):
            if val == 1:
                ret = itf.SetPointData(mkr_idx, 3, start_fr+idx, var_const) 
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
        
def set_analog_data(itf, sig_name, sig_value, start_frame=None, end_frame=None, log=False):
    """
    Update the value of an analog channel.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    sig_value : numpy array
        A new analog channel value. This is assumed as a scaled one.        
    start_frame: None or int, optional
        User-defined start frame. The default is None.
    end_frame: None or int, optional
        User-defined end frame. The default is None. 
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        fr_check, start_fr, end_fr = check_frame_range_valid(itf, start_frame, end_frame, log=log)
        if not fr_check:
            err_msg = 'Given "start_frame" is not proper'
            raise ValueError(err_msg)
        n_frs = end_fr-start_fr+1
        av_ratio = get_analog_video_ratio(itf, log=log)
        if sig_value.ndim != 1 or sig_value.shape[0] != (n_frs*av_ratio):
            err_msg = 'The dimension of the input is not compatible'
            raise ValueError(err_msg)
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            err_msg = f'Unable to get the index of "{sig_name}"'
            raise ValueError(err_msg)
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_scale = get_analog_scale(itf, sig_name, log=log)
        sig_offset = np.float32(get_analog_offset(itf, sig_name, log=log))
        sig_value_unscaled = np.asarray(sig_value, dtype=np.float32)/(sig_scale*gen_scale)+sig_offset
        variant = win32.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R4, sig_value_unscaled.tolist())
        ret = itf.SetAnalogDataEx(sig_idx, start_fr, variant)
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
        
def set_analog_subframe_data(itf, sig_name, sig_value, start_frame, sub_frame, log=False):
    """
    Update the value of analog channel subframe data.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    sig_name : str
        Analog channel name.
    sig_value : float
        A new analog channel subframe value.
    start_frame : int
        Start frame number.
    sub_frame : int
        Sub frame number.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    bool
        True or False.

    """
    try:
        fr_check, start_fr, _ = check_frame_range_valid(itf, start_frame, None, log=log)
        if not fr_check:
            err_msg = 'Given "start_frame" is not proper'
            raise ValueError(err_msg)
        av_ratio = get_analog_video_ratio(itf, log=log)
        if sub_frame < 1 or sub_frame > av_ratio:
            err_msg = f'"sub_frame" should be between 1 and {av_ratio}'
            raise ValueError(err_msg)
        sig_idx = get_analog_index(itf, sig_name, log=log)
        if sig_idx == -1 or sig_idx is None:
            err_msg = f'Unable to get the index of "{sig_name}"'
            raise ValueError(err_msg)
        gen_scale = get_analog_gen_scale(itf, log=log)
        sig_scale = get_analog_scale(itf, sig_name, log=log)
        sig_offset = np.float32(get_analog_offset(itf, sig_name, log=log))
        sig_value_unscaled = np.float32(sig_value)/(sig_scale*gen_scale)+sig_offset
        variant = win32.VARIANT(pythoncom.VT_R4, sig_value_unscaled)
        ret = itf.SetAnalogData(sig_idx, start_fr, sub_frame, variant)
        return [False, True][ret]
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise        

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
    try:
        if log: logger.debug(f'Start recovery of "{tgt_mkr_name}" ...')
        n_total_frs = get_num_frames(itf, log=log)
        tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
        if tgt_mkr_data is None:
            err_msg = f'Unable to get the information of "{tgt_mkr_name}"'
            raise ValueError(err_msg)            
        tgt_mkr_coords = tgt_mkr_data[:,0:3]
        tgt_mkr_resid = tgt_mkr_data[:,3]
        tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
        n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
        if n_tgt_mkr_valid_frs == 0:
            if log: logger.info(f'Recovery of "{tgt_mkr_name}" skipped: no valid target marker frame')
            return False, n_tgt_mkr_valid_frs
        if n_tgt_mkr_valid_frs == n_total_frs:
            if log: logger.info(f'Recovery of "{tgt_mkr_name}" skipped: all target marker frames valid')
            return False, n_tgt_mkr_valid_frs
        dict_cl_mkr_coords = {}
        dict_cl_mkr_valid = {}
        cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
        for mkr in cl_mkr_names:
            mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
            if mkr_data is None:
                err_msg = f'Unable to get the information of "{mkr}"'
                raise ValueError(err_msg)
            dict_cl_mkr_coords[mkr] = mkr_data[:, 0:3]
            dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
            cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
        all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
        if not np.any(all_mkr_valid_mask):
            if log: logger.info(f'Recovery of "{tgt_mkr_name}" skipped: no common valid frame among markers')
            return False, n_tgt_mkr_valid_frs
        cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
        if not np.any(cl_mkr_only_valid_mask):
            if log: logger.info(f'Recovery of "{tgt_mkr_name}" skipped: cluster markers not helpful')
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
        mat_rot = np.asarray([vec_x.T, vec_y.T, vec_z.T]).T
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
        set_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, None, log=log)
        set_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, None, log=log)
        n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
        if log: logger.info(f'Recovery of "{tgt_mkr_name}" finished')
        return True, n_tgt_mkr_valid_frs_updated
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise    
    except ValueError as err:
        if log: logger.error(err)
        raise
        
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
    try:
        if log: logger.debug(f'Start recovery of "{tgt_mkr_name}" ...')
        n_total_frs = get_num_frames(itf, log=log)
        tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
        if tgt_mkr_data is None:
            err_msg = f'Unable to get the information of "{tgt_mkr_name}"'
            raise ValueError(err_msg)        
        tgt_mkr_coords = tgt_mkr_data[:,0:3]
        tgt_mkr_resid = tgt_mkr_data[:,3]
        tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
        n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
        if n_tgt_mkr_valid_frs == 0:
            if log: logger.info(f'Recovery of "{tgt_mkr_name}" skipped: no valid target marker frame')
            return False, n_tgt_mkr_valid_frs
        if n_tgt_mkr_valid_frs == n_total_frs:
            if log: logger.info('Recovery of "{tgt_mkr_name}" skipped: all target marker frames valid')
            return False, n_tgt_mkr_valid_frs    
        dict_cl_mkr_coords = {}
        dict_cl_mkr_valid = {}
        cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
        for mkr in cl_mkr_names:
            mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
            if mkr_data is None:
                err_msg = f'Unable to get the information of "{mkr}"'
                raise ValueError(err_msg)
            dict_cl_mkr_coords[mkr] = mkr_data[:,0:3]
            dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
            cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
        all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
        if not np.any(all_mkr_valid_mask):
            if log: logger.info('Recovery of "{tgt_mkr_name}" skipped: no common valid frame among markers')
            return False, n_tgt_mkr_valid_frs
        cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
        if not np.any(cl_mkr_only_valid_mask):
            if log: logger.info('Recovery of "{tgt_mkr_name}" skipped: cluster markers not helpful')
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
        mat_rot = np.asarray([vec_x.T, vec_y.T, vec_z.T]).T
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
        set_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, None, log=log)
        set_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, None, log=log)
        n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
        if log: logger.info(f'Recovery of "{tgt_mkr_name}" finished')
        return True, n_tgt_mkr_valid_frs_updated
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise        

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
    try:
        if log: logger.debug(f'Start gap filling of "{tgt_mkr_name}" ...')     
        n_total_frs = get_num_frames(itf, log=log)
        tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
        if tgt_mkr_data is None:
            err_msg = f'Unable to get the information of "{tgt_mkr_name}"'
            raise ValueError(err_msg)        
        tgt_mkr_coords = tgt_mkr_data[:,0:3]
        tgt_mkr_resid = tgt_mkr_data[:,3]
        tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
        n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
        if n_tgt_mkr_valid_frs == 0:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no valid target marker frame')
            return False, n_tgt_mkr_valid_frs
        if n_tgt_mkr_valid_frs == n_total_frs:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: all target marker frames valid')
            return False , n_tgt_mkr_valid_frs   
        dict_cl_mkr_coords = {}
        dict_cl_mkr_valid = {}
        cl_mkr_valid_mask = np.ones((n_total_frs), dtype=bool)
        for mkr in cl_mkr_names:
            mkr_data = get_marker_data(itf, mkr, blocked_nan=False, log=log)
            if mkr_data is None:
                err_msg = f'Unable to get the information of "{mkr}"'
                raise ValueError(err_msg)            
            dict_cl_mkr_coords[mkr] = mkr_data[:,0:3]
            dict_cl_mkr_valid[mkr] = np.where(np.isclose(mkr_data[:,3], -1), False, True)
            cl_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, dict_cl_mkr_valid[mkr])
        all_mkr_valid_mask = np.logical_and(cl_mkr_valid_mask, tgt_mkr_valid_mask)
        if not np.any(all_mkr_valid_mask):
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no common valid frame among markers')
            return False, n_tgt_mkr_valid_frs
        cl_mkr_only_valid_mask = np.logical_and(cl_mkr_valid_mask, np.logical_not(tgt_mkr_valid_mask))
        if not np.any(cl_mkr_only_valid_mask):
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: cluster markers not helpful')
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
            set_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, None, log=log)
            set_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, None, log=log)
            n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" finished')
            return True, n_tgt_mkr_valid_frs_updated
        else:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped')
            return False, n_tgt_mkr_valid_frs
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise        

def fill_marker_gap_pattern(itf, tgt_mkr_name, dnr_mkr_name, search_span_offset=5, min_needed_frs=10, log=False):
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
    search_span_offset : int, optional
        Offset for backward and forward search spans. The default is 5.
    min_needed_frs : int, optional
        Minimum required valid frames in both search spans. The default is 10.        
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
    This function is adapted from 'fill_marker_gap_pattern()' function in the GapFill module, see [1] in the References.   
    
    References
    ----------
    .. [1] https://github.com/mkjung99/gapfill
    
    """
    try:
        if log: logger.debug(f'Start gap filling of "{tgt_mkr_name}" ...')    
        n_total_frs = get_num_frames(itf, log=log)
        tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
        if tgt_mkr_data is None:
            err_msg = f'Unable to get the information of "{tgt_mkr_name}"'
            raise ValueError(err_msg)
        tgt_mkr_coords = tgt_mkr_data[:, 0:3]
        tgt_mkr_resid = tgt_mkr_data[:, 3]
        tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
        n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)
        if n_tgt_mkr_valid_frs == 0:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no valid target marker frame')
            return False, n_tgt_mkr_valid_frs
        if n_tgt_mkr_valid_frs == n_total_frs:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: all target marker frames valid')
            return False , n_tgt_mkr_valid_frs    
        dnr_mkr_data = get_marker_data(itf, dnr_mkr_name, blocked_nan=False, log=log)
        if dnr_mkr_data is None:
            err_msg = f'Unable to get the information of "{dnr_mkr_name}"'
            raise ValueError(err_msg)        
        dnr_mkr_coords = dnr_mkr_data[:, 0:3]
        dnr_mkr_resid = dnr_mkr_data[:, 3]
        dnr_mkr_valid_mask = np.where(np.isclose(dnr_mkr_resid, -1), False, True)
        if not np.any(dnr_mkr_valid_mask):
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no valid donor marker frame')
            return False, n_tgt_mkr_valid_frs    
        both_mkr_valid_mask = np.logical_and(tgt_mkr_valid_mask, dnr_mkr_valid_mask)
        if not np.any(both_mkr_valid_mask):
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no valid common frame between target and donor markers')
            return False, n_tgt_mkr_valid_frs        
        tgt_mkr_invalid_frs = np.where(~tgt_mkr_valid_mask)[0]
        tgt_mkr_invalid_gaps = np.split(tgt_mkr_invalid_frs, np.where(np.diff(tgt_mkr_invalid_frs)!=1)[0]+1)
        b_updated = False
        for gap in tgt_mkr_invalid_gaps:
            # Skip if gap size is zero
            if gap.size == 0: continue
            # Skip if gap is either at the first or at the end of the entire frames.
            if gap.min()==0 or gap.max()==n_total_frs-1: continue
            search_span = np.int(np.ceil(gap.size/2))+search_span_offset
            gap_near_tgt_mkr_valid_mask = np.zeros((n_total_frs,), dtype=bool)
            for i in range(gap.min()-1, gap.min()-1-search_span, -1):
                if i >= 0: gap_near_tgt_mkr_valid_mask[i]=True
            for i in range(gap.max()+1, gap.max()+1+search_span, 1):
                if i < n_total_frs: gap_near_tgt_mkr_valid_mask[i]=True
            gap_near_tgt_mkr_valid_mask = np.logical_and(gap_near_tgt_mkr_valid_mask, tgt_mkr_valid_mask)
            # Skip if total number of available target marker frames near the gap within search span is less then minimum required number.
            if np.sum(gap_near_tgt_mkr_valid_mask) < min_needed_frs: continue
            # Skip if there is any invalid frame of the donor marker during the gap period.
            if np.any(~dnr_mkr_valid_mask[gap]): continue
            gap_near_both_mkr_valid_mask = np.logical_and(gap_near_tgt_mkr_valid_mask, dnr_mkr_valid_mask)
            gap_near_both_mkr_valid_frs = np.where(gap_near_both_mkr_valid_mask)[0]
            for idx, fr in np.ndenumerate(gap):
                search_idx = np.searchsorted(gap_near_both_mkr_valid_frs, fr)
                if search_idx == 0:
                    fr0 = gap_near_both_mkr_valid_frs[0]
                    fr1 = gap_near_both_mkr_valid_frs[1]
                elif search_idx >= gap_near_both_mkr_valid_frs.shape[0]:
                    fr0 = gap_near_both_mkr_valid_frs[gap_near_both_mkr_valid_frs.shape[0]-2]
                    fr1 = gap_near_both_mkr_valid_frs[gap_near_both_mkr_valid_frs.shape[0]-1]
                else:
                    fr0 = gap_near_both_mkr_valid_frs[search_idx-1]
                    fr1 = gap_near_both_mkr_valid_frs[search_idx]
                # Skip if the target marker frame fr is outside of range.
                if fr <= fr0 or fr >= fr1: continue
                # Skip if the donor marker is invalid at either fr0 or fr1.
                if ~dnr_mkr_valid_mask[fr0] or ~dnr_mkr_valid_mask[fr1]: continue
                v_tgt = (tgt_mkr_coords[fr1]-tgt_mkr_coords[fr0])*np.float32(fr-fr0)/np.float32(fr1-fr0)+tgt_mkr_coords[fr0]
                v_dnr = (dnr_mkr_coords[fr1]-dnr_mkr_coords[fr0])*np.float32(fr-fr0)/np.float32(fr1-fr0)+dnr_mkr_coords[fr0]
                tgt_mkr_coords[fr] = v_tgt-v_dnr+dnr_mkr_coords[fr]
                tgt_mkr_resid[fr] = 0.0
                b_updated = True
        if b_updated:
            set_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, None, log=log)
            set_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, None, log=log)
            n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" finished')
            return True, n_tgt_mkr_valid_frs_updated
        else:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped')
            return False, n_tgt_mkr_valid_frs
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise

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
    try:
        if log: logger.debug(f'Start gap filling of "{tgt_mkr_name}" ...')
        n_total_frs = get_num_frames(itf, log=log)
        tgt_mkr_data = get_marker_data(itf, tgt_mkr_name, blocked_nan=False, log=log)
        if tgt_mkr_data is None:
            err_msg = f'Unable to get the information of "{tgt_mkr_name}"'
            raise ValueError(err_msg)        
        tgt_mkr_coords = tgt_mkr_data[:, 0:3]
        tgt_mkr_resid = tgt_mkr_data[:, 3]
        tgt_mkr_valid_mask = np.where(np.isclose(tgt_mkr_resid, -1), False, True)
        n_tgt_mkr_valid_frs = np.count_nonzero(tgt_mkr_valid_mask)    
        if n_tgt_mkr_valid_frs == 0:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: no valid target marker frame')
            return False, n_tgt_mkr_valid_frs
        if n_tgt_mkr_valid_frs == n_total_frs:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped: all target marker frames valid')
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
            set_marker_pos(itf, tgt_mkr_name, tgt_mkr_coords, None, None, log=log)
            set_marker_resid(itf, tgt_mkr_name, tgt_mkr_resid, None, None, log=log)
            n_tgt_mkr_valid_frs_updated = np.count_nonzero(np.where(np.isclose(tgt_mkr_resid, -1), False, True))
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" finished')
            return True, n_tgt_mkr_valid_frs_updated
        else:
            if log: logger.info(f'Gap filling of "{tgt_mkr_name}" skipped')
            return False, n_tgt_mkr_valid_frs
    except pythoncom.com_error as err:
        if log: logger.error(err.excepinfo[2])
        raise
    except ValueError as err:
        if log: logger.error(err)
        raise
    
def export_trc(itf, f_path, rot_mat=np.eye(3), filt_fc=None, filt_order=2, tgt_mkr_names=None, start_fr=None, end_fr=None, fmt='%.6f', log=False):
    """
    Export a TRC format file, which is compatible with OpenSim for markers.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    f_path : str
        Path of the output TRC file to export.
    rot_mat : list or ndarray
        Transformation matrix between the original lab coordinate system and the desired lab coordinate system for export.
    filt_fc : float, optional
        Cut-off frequency of the zero-lag low-pass butterworth filter. The default is None.
    filt_order : int, optional
        Order of the zero-lag low-pass butterworth filter. The default is 2.
    tgt_mkr_names : list, optional
        Specific target marker names. The default is None.
    start_fr : int, optional
        Start frame for export. The default is None.
    end_fr : int, optional
        End frame for export. The default is None.
    fmt : str, optional
        A single format string for all marker trajectories.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    None
        None.
    
    """    
    dict_pts = get_dict_markers(itf, log=log)
    orig_start_fr = 1
    vid_fps = int(get_video_fps(itf, log=log))
    if start_fr is None:
        start_fr = get_first_frame(itf, log=log)
    if end_fr is None:
        end_fr = get_last_frame(itf, log=log)
    n_vid_frs = end_fr-start_fr+1
    vid_frs = np.linspace(start_fr, end_fr, n_vid_frs, dtype=np.int32)
    vid_times = (vid_frs-orig_start_fr)/vid_fps
    orig_vid_frs = get_video_frames(itf, log=log)
    frs_sel_mask = np.logical_and(orig_vid_frs>=vid_frs[0], orig_vid_frs<=vid_frs[-1])
    total_mkr_names = list(dict_pts['LABELS'])
    if tgt_mkr_names is None:
        mkr_names = total_mkr_names
    else:
        mkr_names = [x for x in list(tgt_mkr_names) if x in total_mkr_names]
    mkr_unit = 'mm'
    hdr_row0 = f'PathFileType\t4\t(X/Y/Z)\t{os.path.basename(f_path)}'
    hdr_row1 = 'DataRate\tCameraRate\tNumFrames\tNumMarkers\tUnits\tOrigDataRate\tOrigDataStartFrame\tOrigNumFrames'
    hdr_row2 = f'{vid_fps}\t{vid_fps}\t{n_vid_frs}\t{len(mkr_names)}\t{mkr_unit}\t{vid_fps}\t{orig_start_fr}\t{orig_vid_frs.shape[0]}'
    hdr_row3 = 'Frame#\tTime\t'+'\t\t\t'.join(mkr_names)+'\t\t'
    hdr_row4 = '\t\t'+'\t'.join([c+str(n) for n in list(range(1, len(mkr_names)+1)) for c in ['X', 'Y', 'Z']])
    hdr_row5 = '\t'
    hdr_str = '\n'.join([hdr_row0, hdr_row1, hdr_row2, hdr_row3, hdr_row4, hdr_row5])
    output_data = np.zeros((n_vid_frs, 2+3*len(mkr_names)), dtype=float)
    output_data[:,0] = vid_frs
    output_data[:,1] = vid_times
    for mkr_idx, mkr_name in enumerate(mkr_names):
        mkr_pos_raw = np.dot(np.asarray(rot_mat), dict_pts['DATA']['POS'][mkr_name].T).T
        if filt_fc is None:
            mkr_pos = mkr_pos_raw
        else:
            mkr_pos = filt_bw_lp(mkr_pos_raw, filt_fc, vid_fps, order=filt_order)
        output_data[:,2+3*mkr_idx:2+3*mkr_idx+3] = mkr_pos[frs_sel_mask,:]
    fmt_str = f'%d\t{fmt}\t'+'\t'.join([fmt]*(3*len(mkr_names)))
    np.savetxt(f_path, output_data, fmt=fmt_str, delimiter='\t', comments='', header=hdr_str)
    return None

def export_mot(itf, f_path, rot_mat=np.eye(3), filt_fc=None, filt_order=2, threshold=0.0, start_fr=None, end_fr=None, fmt='%.6f', log=False):
    """
    Export a MOT format file, which is compatible with OpenSim for force plates.

    Parameters
    ----------
    itf : win32com.client.CDispatch
        COM object of the C3Dserver.
    f_path : str
        Path of the output MOT file to export.
    rot_mat : list or ndarray
        Transformation matrix between the original lab coordinate system and the desired lab coordinate system for export.
    filt_fc : float, optional
        Cut-off frequency of the zero-lag low-pass butterworth filter. The default is None.
    filt_order : int, optional
        Order of the zero-lag low-pass butterworth filter. The default is 2.
    threshold : float, optional
        Threshold value of Fz (force plate local) to determine the frames where all forces and moments will be zero.
    start_fr : int, optional
        Start frame for export. The default is None.
    end_fr : int, optional
        End frame for export. The default is None.
    fmt : str, optional
        A single format string for all force plate related values.
    log : bool, optional
        Whether to write logs or not. The default is False.

    Returns
    -------
    None
        None.
    
    """      
    orig_start_fr = 1
    if start_fr is None:
        start_fr = get_first_frame(itf, log=log)
    if end_fr is None:
        end_fr = get_last_frame(itf, log=log) 
    vid_fps = int(get_video_fps(itf, log=log))
    anal_fps = get_analog_fps(itf, log=log)
    av_ratio = get_analog_video_ratio(itf, log=log)
    n_vid_frs = end_fr-start_fr+1
    vid_frs = np.linspace(start_fr, end_fr, n_vid_frs, dtype=np.int32)
    vid_times = (vid_frs-orig_start_fr)/vid_fps
    first_fr = get_first_frame(itf, log=log)
    last_fr = get_last_frame(itf, log=log)
    anal_start_t = np.float32(first_fr-orig_start_fr)/vid_fps
    anal_end_t = np.float32(last_fr-orig_start_fr)/vid_fps+np.float32(av_ratio-1)/anal_fps
    anal_steps = (last_fr-first_fr+1)*av_ratio
    anal_times = np.linspace(start=anal_start_t, stop=anal_end_t, num=anal_steps, dtype=np.float32)
    anal_sel_mask = np.logical_and(anal_times>vid_times[0]-(1.0/anal_fps), anal_times<vid_times[-1]+(1.0/anal_fps))
    anal_sel_times = anal_times[anal_sel_mask]
    fp_output = get_fp_output(itf, threshold=threshold, filt_fc=filt_fc, filt_order=filt_order, log=log)
    cnt_fp = len(fp_output)
    output_col_names = []
    for i in range(cnt_fp):
        output_col_names.append(f'ground_force{i+1}_vx')
        output_col_names.append(f'ground_force{i+1}_vy')
        output_col_names.append(f'ground_force{i+1}_vz')        
        output_col_names.append(f'ground_force{i+1}_px')
        output_col_names.append(f'ground_force{i+1}_py')
        output_col_names.append(f'ground_force{i+1}_pz')
        output_col_names.append(f'ground_torque{i+1}_x')
        output_col_names.append(f'ground_torque{i+1}_y')
        output_col_names.append(f'ground_torque{i+1}_z')
    hdr_row0 = f'{os.path.basename(f_path)}'
    hdr_row1 = f'datacolumns\t{1+9*cnt_fp}'
    hdr_row2 = f'datarows\t{anal_sel_times.shape[0]}'
    hdr_row3 = f'range\t{anal_sel_times[0]:{fmt.replace("%","")}}\t{anal_sel_times[-1]:{fmt.replace("%","")}}'
    hdr_row4 = 'endheader'
    hdr_row5 = 'time\t'+'\t'.join(output_col_names)
    hdr_str = '\n'.join([hdr_row0, hdr_row1, hdr_row2, hdr_row3, hdr_row4, hdr_row5])
    output_data = np.zeros((anal_sel_times.shape[0], len(output_col_names)+1), dtype=float)
    output_data[:,0] = anal_sel_times
    for i in range(cnt_fp):
        cop_lab = fp_output[i]['COP_LAB'][anal_sel_mask,:]
        f_cop_lab = fp_output[i]['F_COP_LAB'][anal_sel_mask,:]
        m_cop_lab = fp_output[i]['M_COP_LAB'][anal_sel_mask,:]
        output_data[:,9*i+1:9*i+4] = np.dot(np.asarray(rot_mat), f_cop_lab.T).T
        output_data[:,9*i+4:9*i+7] = np.dot(np.asarray(rot_mat), cop_lab.T).T
        output_data[:,9*i+7:9*i+10] = np.dot(np.asarray(rot_mat), m_cop_lab.T).T
    np.savetxt(f_path, output_data, fmt=fmt, delimiter='\t', comments='', header=hdr_str)
    return None