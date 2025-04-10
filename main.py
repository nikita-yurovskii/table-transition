from docx import Document
import excel_table
import merged
import make_width_perfect
import transition
def create_table(document, rows, cols, style='Table Grid'):
    """
    Создает таблицу в документе Word заданного размера и стиля.

    Args:
        document: Объект документа Word (docx.Document).
        rows: Количество строк в таблице.
        cols: Количество столбцов в таблице.
        style: Стиль таблицы (необязательный, по умолчанию 'Table Grid').
               Смотрите в документации python-docx доступные стили:
               https://python-docx.readthedocs.io/en/latest/api/table.html#docx.enum.style.WD_STYLE_TYPE
    """

    table = document.add_table(rows=rows, cols=cols, style=style)
    return table  # Возвращаем объект таблицы для дальнейшей работы


if __name__ == '__main__':
    # Пример использования
    file = "input.xlsx"
    sheet = "Общие сведения"
    document = Document()
    mr, mc = excel_table.get_print_area_size(file, sheet)
    print(mr,mc)
    pogr_r = excel_table.get_first_visible_row_index(file,sheet)
    pogr_c = excel_table.get_first_visible_column_index(file, sheet)+3
    table1 = create_table(document, mr, mc)
    table_width_map = excel_table.get_cell_widths_in_range(file, sheet, pogr_r, pogr_c)[0]
    table_heights_map = excel_table.get_cell_widths_in_range(file, sheet, pogr_r, pogr_c)[1]
    make_width_perfect.set_cell_widths_in_table(table1, table_width_map, mr, mc)
    #make_width_perfect.set_cell_height_in_table(table1, table_heights_map, mr, mc)






    transition.transfer_excel_to_word_table(file, sheet, table1, excel_table.get_print_area(file, sheet))
    merged.merge_cells_from_coordinates(document, table1, merged.get_merged_cells(file, pogr_r, pogr_c, mr, mc, sheet))
    document.save("output.docx")
    print("Документ 'output.docx' успешно создан с таблицами.")