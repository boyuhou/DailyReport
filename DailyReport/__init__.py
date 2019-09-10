import os
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
from decimal import Decimal, ROUND_HALF_UP
import math
import pandas as pd
import numpy as np
from PIL import Image
import comtypes.client

WORD2PDF_FORMAT = 17


class COLOURS:
    COLOUR_BLACK = '000000'
    COLOUR_RED = 'FF0000'
    COLOUR_WHITE = 'FFFFFF'
    COLOUR_LIGHT_GRAY = 'D3D3D3'
    COLOUR_ORANGE_YELLOW = 'EBD135'


CELL_FORMAT_DEFAULT_KEY = 'default'
DICT_KEY_ALL = 'all'
DASH_CELL_VALUE = '-'
COLOUR_DEFAULT_FG = COLOURS.COLOUR_BLACK
COLOUR_DEFAULT_BG = COLOURS.COLOUR_WHITE
COLOUR_DEFAULT_HIGHLIGHT = COLOURS.COLOUR_ORANGE_YELLOW

LARGE_CELL_FONT_SIZE = 21
LARGE_ROW_LABEL_FONT_SIZE = 22
LARGE_COL_LABEL_FONT_SIZE = 22

MEDIUM_CELL_FONT_SIZE = 16
MEDIUM_ROW_LABEL_FONT_SIZE = 16
MEDIUM_COL_LABEL_FONT_SIZE = 16

SMALL_CELL_FONT_SIZE = 12
SMALL_ROW_LABEL_FONT_SIZE = 12
SMALL_COL_LABEL_FONT_SIZE = 12


def cell_bg_shade(row_group, col_group, cell, highlight_thresholds, cur_row):
    if row_group is None or col_group is None:
        return COLOUR_DEFAULT_BG
    elif row_group == col_group:
        return COLOURS.COLOUR_LIGHT_GRAY
    else:
        return COLOUR_DEFAULT_BG


def cell_bg_asset_threshold_highlight_or_shade(row_group, col_group, cell, highlight_thresholds, cur_row):
    if not highlight_thresholds:
        return cell_bg_shade(row_group, col_group, cell, highlight_thresholds, cur_row)

    if row_group in highlight_thresholds and highlight_thresholds[row_group] <= abs(cell):
        return COLOUR_DEFAULT_HIGHLIGHT

    if col_group in highlight_thresholds and highlight_thresholds[col_group] <= abs(cell):
        return COLOUR_DEFAULT_HIGHLIGHT

    return cell_bg_shade(row_group, col_group, cell, highlight_thresholds, cur_row)


def cell_bg_every_other_row(row_group, col_group, cell, highlight_thresholds, cur_row):
    if cur_row % 2 == 0:
        return COLOURS.COLOUR_LIGHT_GRAY


def cell_fg_colour_by_column_name(column_name, colour_dict):
    return colour_dict.get(column_name, None)


DEFAULT_GRIDS_CELL_FORMAT_CONFIG = {
    'default': {
        'cell_font_size': MEDIUM_CELL_FONT_SIZE,
        'row_label_font_size': MEDIUM_ROW_LABEL_FONT_SIZE,
        'col_label_font_size': MEDIUM_COL_LABEL_FONT_SIZE,
        'make_labels_bold': False,
        'format_num_decimal': 2,
        'background_colour_ftn': cell_bg_every_other_row,
        'highlight_thresholds': None
    }
}


def combine_images(list_im, target_path):
    imgs = [Image.open(i) for i in list_im]
    # pick the image which is the smallest, and resize the others to match it (can be arbitrary image shape here)
    min_shape = sorted([(np.sum(i.size), i.size) for i in imgs])[0][1]
    imgs_comb = np.hstack([np.asarray(i.resize(min_shape)) for i in imgs])

    # save that beautiful picture
    imgs_comb = Image.fromarray(imgs_comb)
    imgs_comb.save(target_path)


def save_doc(context, inline_imgs, file_dict, template_path, result_path):
    tpl = DocxTemplate(template_path)
    for cur_key in file_dict:
        tpl.replace_pic(cur_key, file_dict[cur_key])
    inline_image_context = dict()
    for key in inline_imgs:
        width = 170
        width_key = key + '_width_percent'
        if width_key in inline_imgs:
            width = inline_imgs[width_key] * width
        inline_image_context[key] = InlineImage(tpl, inline_imgs[key], width=Mm(width))
    context = {**context, **inline_image_context}
    tpl.render(context)
    tpl.save(result_path)


def doc2pdf(input_path, output_path):
    word = comtypes.client.CreateObject('Word.Application')
    try:
        doc = word.Documents.Open(input_path)
        try:
            doc.ExportAsFixedFormat(OutputFileName=output_path,
                                    ExportFormat=WORD2PDF_FORMAT,  # 17 = PDF output, 18=XPS output
                                    OpenAfterExport=False,
                                    OptimizeFor=0,  # 0=Print (higher res), 1=Screen (lower res)
                                    CreateBookmarks=1,
                                    # 0=No bookmarks, 1=Heading bookmarks only, 2=bookmarks match word bookmarks
                                    DocStructureTags=True
                                    )
        finally:
            doc.Close()
    finally:
        word.Quit()


def get_group_value(group_path, instrument_list, indicator_name):
    result = {}
    for i in instrument_list:
        _path = os.path.join(group_path, i + '.csv')
        _df = pd.read_csv(_path, header=None)
        result[i] = _df.iloc[-1, -1]
    return pd.DataFrame(result, [indicator_name])


def get_specific_value(group_path, instrument, column_name):
    _path = os.path.join(group_path, instrument + '.csv')
    _df = pd.read_csv(_path)
    return _df.loc[:, column_name].iloc[-1]


def get_value_area_style(val, vah, vpoc, atr):
    style = ''
    if vah-val >= 0.5 * atr:
        style = style + 'Wide, '
    if vah - vpoc <= 0.1 * atr:
        style = style + 'Stacked Up, '
    if vpoc - val <= 0.1 * atr:
        style = style + 'Stacked Bottom, '
    if vah - val <= 0.25 * atr:
        style = style + 'Tight, '
    if style == '':
        style = 'Normal, '
    return style[:-2]


class GridBuilder:

    @staticmethod
    def __round_number(num, num_digits):
        decimal_offset = '0' if num_digits == 0 else str(math.pow(10, -1 * num_digits))
        return Decimal(str(num)).quantize(Decimal(decimal_offset), rounding=ROUND_HALF_UP)

    @staticmethod
    def __get_cell_info_dict(cell,
                             format_percent,
                             format_num_decimal,
                             convert_to_millions,
                             make_zero_dash,
                             make_nan_dash,
                             make_negatives_red,
                             make_cell_bold,
                             foreground_colour=None,
                             background_colour=None,
                             add_type_suffix=None,
                             font_size=None):
        cell_fg_colour = COLOUR_DEFAULT_FG if foreground_colour is None else foreground_colour
        cell_bg_colour = COLOUR_DEFAULT_BG if background_colour is None else background_colour

        cell_is_a_number = False
        cell_value = str(cell)

        if make_nan_dash and np.isnan(cell):
            cell_value = DASH_CELL_VALUE
        elif type(cell) == float or type(cell) == np.float64 or type(cell) == np.int64:
            if make_zero_dash and cell == 0.00:
                cell_value = DASH_CELL_VALUE
            else:
                cell_is_a_number = True
                updated_cell_value = cell
                format_pattern = '{}'

                if format_percent:
                    updated_cell_value = cell * 100
                    if format_num_decimal is not None:
                        updated_cell_value = GridBuilder.__round_number(updated_cell_value, format_num_decimal)

                    if add_type_suffix:
                        format_pattern += '%'
                elif convert_to_millions:
                    updated_cell_value = cell / 1000000
                    if format_num_decimal is not None:
                        updated_cell_value = GridBuilder.__round_number(updated_cell_value, format_num_decimal)

                    format_pattern = '{0:,f}' if format_num_decimal is not None else '{}'
                    if add_type_suffix:
                        format_pattern += 'M'
                else:
                    if format_num_decimal is not None:
                        updated_cell_value = GridBuilder.__round_number(updated_cell_value, format_num_decimal)

                    format_pattern = '{0:,f}' if format_num_decimal is not None else '{}'

                cell_value = format_pattern.format(updated_cell_value)

        # Text Colour
        if make_negatives_red and cell_is_a_number and cell < 0:
            cell_fg_colour = COLOURS.COLOUR_RED

        return {
            'value': RichText(cell_value, color=cell_fg_colour, bold=make_cell_bold, size=font_size),
            'bg': cell_bg_colour
        }

    @staticmethod
    def __get_df_row_segment(df, row_labels):
        result = df.loc[row_labels]
        return result if isinstance(result, pd.DataFrame) else result.to_frame().T

    @staticmethod
    def __get_df_col_segment(df, col_labels):
        return df[col_labels]

    """
    Given a row in df, return the corresponding RichText object for each cell
    """

    @staticmethod
    def __get_column_structure(df_row, cell_format_config, cur_row, row_group=None, col_group=None):
        result_cols = []
        column_labels = df_row.columns.tolist()
        columns = df_row.values[0]

        for i in range(len(columns)):
            cell_config = cell_format_config if column_labels[i] not in cell_format_config else cell_format_config[
                column_labels[i]]

            background_colour = None
            if cell_config.get('background_colour_ftn', None):
                background_colour = cell_config['background_colour_ftn'](row_group,
                                                                         col_group,
                                                                         columns[i],
                                                                         cell_config['highlight_thresholds'],
                                                                         cur_row
                                                                         )
            text_colour = None
            if cell_config.get('text_colour_ftn', None):
                text_colour = cell_config['text_colour_ftn'](column_labels[i], cell_config['text_colour_dict'])
            result_cols.append(
                GridBuilder.__get_cell_info_dict(
                    cell=columns[i],
                    format_percent=cell_config.get('format_percent', None),
                    format_num_decimal=cell_config.get('format_num_decimal', None),
                    convert_to_millions=cell_config.get('convert_to_millions', None),
                    make_zero_dash=cell_config.get('make_zero_dash', None),
                    make_nan_dash=cell_config.get('make_nan_dash', None),
                    make_negatives_red=cell_config.get('make_negatives_red', None),
                    make_cell_bold=False,
                    foreground_colour=text_colour,
                    background_colour=background_colour,
                    add_type_suffix=cell_config.get('add_type_suffix', None),
                    font_size=cell_config.get('cell_font_size', None)
                )
            )

        return result_cols

    @staticmethod
    def __merge_formatted_multiindex_df_into_one(df, section_dict, is_row):
        merged_data = []
        key_set = df.index.get_level_values(0).unique() if is_row else df.columns.get_level_values(0).unique()
        for cur_section in key_set:
            merged_data = merged_data + section_dict[cur_section]

        return merged_data

    @staticmethod
    def __get_column_structure_by_section(df, cell_format_config, cur_row, row_group=None):
        is_columns_multiindex = isinstance(df.columns, pd.MultiIndex)
        col_dict = {}

        if is_columns_multiindex:
            column_level_sections = df.columns.get_level_values(0).unique()

            for cur_section in column_level_sections:
                cell_section_config = cell_format_config[
                    CELL_FORMAT_DEFAULT_KEY] if cur_section not in cell_format_config else cell_format_config[
                    cur_section]
                col_dict[cur_section] = GridBuilder.__get_column_structure(
                    GridBuilder.__get_df_col_segment(df, cur_section), cell_section_config,
                    cur_row, row_group, cur_section)

            col_dict[DICT_KEY_ALL] = GridBuilder.__merge_formatted_multiindex_df_into_one(df, col_dict, is_row=False)
        else:
            if len(cell_format_config) > 1:
                for column in df.columns:
                    cell_section_config = cell_format_config[
                        CELL_FORMAT_DEFAULT_KEY] if column not in cell_format_config else cell_format_config[
                        column]
                    col_dict[column] = GridBuilder.__get_column_structure(df[[column]], cell_section_config, cur_row)
                col_dict[DICT_KEY_ALL] = GridBuilder.__merge_formatted_multiindex_df_into_one(df, col_dict,
                                                                                              is_row=False)
            else:
                col_dict[DICT_KEY_ALL] = GridBuilder.__get_column_structure(df,
                                                                            cell_format_config[CELL_FORMAT_DEFAULT_KEY],
                                                                            cur_row)

        return col_dict

    @staticmethod
    def __get_row_structure(df, cell_format_config, row_group=None):
        row_list = []
        row_id = 0
        for cur_row_data in df.index:
            row_id = row_id + 1
            col_dict = GridBuilder.__get_column_structure_by_section(GridBuilder.__get_df_row_segment(df, cur_row_data),
                                                                     cell_format_config, row_id, row_group)
            col_dict['label'] = RichText(cur_row_data,
                                         bold=cell_format_config[CELL_FORMAT_DEFAULT_KEY].get('make_labels_bold',
                                                                                              False),
                                         size=cell_format_config[CELL_FORMAT_DEFAULT_KEY].get('row_label_font_size', 8))
            # if background colour is specified:
            if (CELL_FORMAT_DEFAULT_KEY in cell_format_config) & (
                    'background_colour_ftn' in cell_format_config[CELL_FORMAT_DEFAULT_KEY]):
                background_colour = cell_format_config[CELL_FORMAT_DEFAULT_KEY]['background_colour_ftn'](None, None,
                                                                                                         None, None,
                                                                                                         row_id)
                col_dict['bg'] = background_colour

            row_list.append(col_dict)

        return row_list

    @staticmethod
    def __build_structure(df, cell_format_config):
        data_dict = {}
        column_dict = {}

        # Get Data
        is_row_multiindex = isinstance(df.index, pd.MultiIndex)
        if is_row_multiindex:
            row_level_sections = df.index.get_level_values(0).unique().tolist()

            for cur_section in row_level_sections:
                data_dict[cur_section] = GridBuilder.__get_row_structure(
                    GridBuilder.__get_df_row_segment(df, cur_section), cell_format_config,
                    cur_section)

            data_dict[DICT_KEY_ALL] = GridBuilder.__merge_formatted_multiindex_df_into_one(df, data_dict, is_row=True)
        else:
            data_dict[DICT_KEY_ALL] = GridBuilder.__get_row_structure(df, cell_format_config)

        # Get Columns
        is_columns_multiindex = isinstance(df.columns, pd.MultiIndex)
        if is_columns_multiindex:
            column_level_sections = df.columns.get_level_values(0).unique()

            for cur_section in column_level_sections:
                column_dict[cur_section] = GridBuilder.__get_rich_text_columns(
                    GridBuilder.__get_df_col_segment(df, cur_section).columns.tolist(),
                    cell_format_config)

                column_dict[DICT_KEY_ALL] = GridBuilder.__get_rich_text_columns(df.columns.get_level_values(1).tolist(),
                                                                                cell_format_config)
        else:
            column_dict[DICT_KEY_ALL] = GridBuilder.__get_rich_text_columns(df.columns.tolist(), cell_format_config)

        return {
            'cols': column_dict,
            'data': data_dict
        }

    @staticmethod
    def __get_rich_text_columns(col_list, cell_format_config):
        ret_list = []
        for cur_col in col_list:
            cur_col = str(cur_col)
            ret_list.append(
                RichText(cur_col, bold=True, size=cell_format_config[CELL_FORMAT_DEFAULT_KEY]['col_label_font_size']))
        return ret_list

    @staticmethod
    def __create_context(structure, prefix):
        columns_key = prefix + '_columns'
        contents_key = prefix + '_contents'

        context = {
            columns_key: structure['cols'],
            contents_key: structure['data']
        }
        return context

    @staticmethod
    def create_grid_context_from_df(df, grid_config, prefix):
        return GridBuilder.__create_context(GridBuilder.__build_structure(df, grid_config), prefix)
