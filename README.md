## Update 2018-07-06

Access to the Activities API has been removed by Microsoft. The retriever tools will no longer pull report data.









## Magic-Unicorn-Tool

This is the beta release of our Office 365 Activities API report parsing tool. It is offered under the BSD License.

## Requirements
    - Python 3.4.3 or above
    - Requests

## Description
    The parsing script is designed to read Microsoft Office 365 Activities API data in csv format and produce a set of
    reports based on the data parsed. All testing for the script was done using Powershell.

## Basic Usage
    Step one - retrieve an activities report using the retriever.py script. This altered version of the original retiever.py
    script is designed to return data in ascending chronological order with the encoding set as utf-8 to avoid any parsing
    errors.

    Step two - run the Magic Unicorn parser using the following command syntax
    > MagicUnicorn_v1.py -i "Microsoft Activities API csv file" -o "Output directory" -t "General report title"

## Reports Generated
    - "General report title"-attachments-activity.tsv
    - "General report title"-search-activity.tsv
    - "General report title"-read-activity-by-time.tsv
    - "General report title"-read-activity-by-item.tsv
    - "General report title"-logon-activity.tsv

## Activities API Data Aquisition
    Modified versions of the CrowdStrike retriever and activity scripts are included in this repository. The MagicUnicorn_v1 parser is designed 
    to work exclusively with the output from these scripts. Directions for use are included in the "CrowdStrike-Retriever-Scripts" folder.
