"""
Author: Michael Bianconi

Allows the user to merge one or more xlsx spreadsheets into one.
"""


import pandas, argparse


def parseArgs() -> dict:
    """
    Parse command line arguments and return a dictionary of the results.
    
    Returns:
    dict: Dictionary of results.
    
    """
    
    parser = argparse.ArgumentParser(description='Merge XLSX files')
    parser.add_argument('-infiles',
                        action='append',
                        nargs='*',
                        metavar=['file','sheets'],
                        help='File path. Optionally, include a list of sheets. ' \
                             'If no sheets are specified, all are loaded.')
    parser.add_argument('-outfile',
                        action='store',
                        nargs='*',
                        help='Path to save the merged file to. Optionally, include a sheet name.')
    parser.add_argument('--include-sheet-name-as-type',
                        action='store_true',
                        default=False,
                        help='Create a column called "Type" that corresponds to the sheet')
    parser.add_argument('--common-columns-only',
                        action='store_true',
                        default=False,
                        help='Ignore columns not shared by all sheets')
    parser.add_argument('--include-index',
                        action='store_true',
                        default=False,
                        help="Add the row number as a column.")
    args = parser.parse_args()

    return {'filenames'             : args.infiles,
            'outpath'               : args.outfile,
            'include_type'          : args.include_sheet_name_as_type,
            'common_columns_only'   : args.common_columns_only,
            'include_index'         : args.include_index}


def loadSheets(filenames: list) -> list:
    """
    Loads all spreadsheets into memory. Load only the specified
    sheets. If no sheets are specified, load them all.
    
    Parameters:
    filesnames (list): First item is the filename. Other items are sheet names.
    
    Returns:
    list: List of DataFrames.
    
    """
    frames = []
    for f in filenames:

        path = f[0]
        sheetnames = f[1:] if len(f) > 1 else None
        
        if sheetnames is None:
            print('LOADING ALL SHEETS FROM ' + path)
            
        else:
            print('LOADING ' + sheetnames.__str__() + ' FROM ' + path)
        
        data = pandas.read_excel(path, sheetnames)
        frames.extend(list(data.items()))
    return frames


def compileColumns(frames: list, add_type: bool) -> list:
    """
    Gets the union of all columns across all frames.
    
    Parameters:
    frames (list):   List of DataFrames.
    add_type (bool): If true, add a column 'Type' to column set.
    
    Returns:
    list: The list of all columns across all sheets.
    """

    columns = set()
    if add_type:
        columns.add('Type')
    for f in frames:
        dataframe = f[1]
        for c in dataframe.columns:
            columns.add(c)
    return list(columns)
   
    
def main():
    """
    Parses command line arguments, loads spreadsheets into memory,
    merges them appropriately, and saves the result to file.
    
    """

    # >>> Load input files into memory <<< #
    args = parseArgs()
    frames = loadSheets(args['filenames'])
    columns = compileColumns(frames, args['include_type'])
    join = 'inner' if args['common_columns_only'] else 'outer'    

    # >>> Create merged spreadsheet <<< #
    outframe = pandas.DataFrame(columns=columns)
    for f in frames:
        if args['include_type']:
            f[1]['Type'] = f[0]
        outframe = pandas.concat((outframe, f[1]), ignore_index=True, sort=True, join=join)
    
    # >>> Save to file <<< #
    outfile_name = args['outpath'][0]
    sheet_name = args['outpath'][1] if len(args['outpath']) > 1 else 'Sheet1'
    print('SAVING TO ' + sheet_name + ' OF ' + args['outpath'][0])
    outframe.to_excel(args['outpath'][0], sheet_name=sheet_name, index=args['include_index'])


if __name__ == '__main__':
    main()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
