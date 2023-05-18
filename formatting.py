"""Convert the JSON to Excel"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

import copy
import re


def list_to_str(l: list):
    """Convert list to comma-separated string"""
    if l:
        return re.sub(r"\[|\]|'", "", str(l))
    else:
        return ""


def profiles_to_dataframe(data: list[dict]) -> pd.DataFrame:
    """Convert the profile JSON data to a dataframe (no videos).

    Args:
        data (list[dict]): The Toughdata follower scraper output
    """
    _data = copy.deepcopy(data)
    for d in _data:
        d.pop("recent_videos")

    return pd.json_normalize(_data)


def videos_to_dataframe(data: list[dict]) -> pd.DataFrame:
    """Convert each profile's videos to a dataframe.

    Args:
        data (list[dict]): The Toughdata follower scraper output
    """
    videos = [v for d in data for v in d["recent_videos"]]
    if videos:
        for v in videos:
            v["categories"] = list_to_str(v["categories"])
            v["video_hashtags"] = list_to_str(v["video_hashtags"])
            v["video_mentions"] = list_to_str(v["video_mentions"])
    return pd.json_normalize(videos)


def data_to_excel(data: list[dict], savename: str | None = None) -> Workbook:
    """Convert the Toughdata output to an Excel workbook.

    Args:
        data (list[dict]): The Toughdata follower scraper output
        savename (str): Path to save the excel sheet to
    """
    # Convert data
    profiles = profiles_to_dataframe(data)
    videos = videos_to_dataframe(data)
    
    # Create workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Profiles" 
    ws2 = wb.create_sheet(title="Videos")
    
    # Add the data
    for r in dataframe_to_rows(profiles, index=False, header=True):
        ws1.append(r)
    for r in dataframe_to_rows(videos, index=False, header=True):
        ws2.append(r)

    # Save
    if savename:
        if not savename.endswith(".xlsx"):
            raise ValueError("savename must end in .xlsx")
        wb.save(savename)

    return wb
