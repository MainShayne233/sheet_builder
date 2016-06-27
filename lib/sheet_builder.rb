module SheetBuilder

  class Cell
    attr_accessor :text
    attr_accessor :title
    attr_accessor :row
    attr_accessor :col
    def initialize attrs
      attrs.each {|key, val| self.send("#{key.to_s}=", val)}
    end
  end

  class Title
    attr_accessor :text
    attr_accessor :comment
    attr_accessor :color
    attr_accessor :list
    attr_accessor :hyperlink

    def initialize attrs
      attrs.each {|key, val| self.send("#{key.to_s}=", val)}
    end

    def combined_style
      combined_style = {b: true, alignment: {horizontal: :center, vertical: :center}, border: {style: :thick, color: '00000000'}}
      combined_style.merge!(bg_color: self.color) if self.color
      combined_style
    end
  end

  class Element
    attr_accessor :text
    attr_accessor :row
    attr_accessor :col
    attr_accessor :color
    attr_accessor :font_color
    attr_accessor :style
    attr_accessor :borders
    attr_accessor :border_thickness
    attr_accessor :merge
    attr_accessor :comment

    def initialize attrs
      attrs.each {|key, val| self.send("#{key.to_s}=", val)}
      self.col -= 1
    end

    def combined_style
      combined_style = {}
      if self.style
        combined_style.merge!(b: true) if self.style.include? :bold
        combined_style.merge!(i: true) if self.style.include? :italic
        combined_style.merge!(u: true) if self.style.include? :underline
        combined_style.merge!(alignment: {horizontal: :center}) if self.style.include? :center
        combined_style.merge!(sz: 12) if self.style.include? :lg_font
      end
      combined_style.merge!(bg_color: self.color) if self.color
      combined_style.merge!(fg_color: self.font_color) if self.font_color
      if self.borders == true
        combined_style.merge! border: {style: self.border_thickness || :medium, color: '00000000'}
      elsif self.borders
        combined_style.merge! border: {style: self.border_thickness || :medium, color: '00000000', edges: self.borders}
      end
      combined_style
    end
  end

  class Blueprint
    attr_accessor :sheet
    attr_accessor :elements
    attr_accessor :column_titles
    attr_accessor :row_titles
    attr_accessor :column_data
    attr_accessor :row_data
    attr_accessor :column_titles_start
    attr_accessor :row_titles_start
    attr_accessor :max
    attr_accessor :column_title_row_height

    def initialize options = []
      self.elements = options[:elements] ? options[:elements].map{|elem| SheetBuilder::Element.new(elem)} : []
      self.column_titles = options[:column_titles] ? options[:column_titles].map{|elem| SheetBuilder::Title.new(elem)} : []
      self.row_titles = options[:row_titles] ? options[:row_titles].map{|elem| SheetBuilder::Title.new(elem)} : []
      self.column_titles_start = options[:column_titles_start] || [1,1]
      self.row_titles_start = options[:row_titles_start] || [1,1]
      self.column_title_row_height = options[:column_title_row_height]
    end

    def build! sheet:, column_data: [], row_data: []

      self.sheet = sheet
      self.column_data = column_data.map{|elem| SheetBuilder::Cell.new(elem)}
      self.row_data = row_data.map{|elem| SheetBuilder::Cell.new(elem)}

      set_column_title_indexes

      set_max_row_and_column

      set_row_title_indexes

      set_column_widths

      set_row_heights

      place_elements

      place_column_titles

      set_lists_for_column_titles

      place_row_titles

      set_lists_for_row_titles

      set_column_title_row_height

      place_column_data

      place_row_data

      move_lists_sheet_to_end

    end

    def set_column_title_indexes
      self.sheet.column_title_indexes = {}
      self.column_titles.each_with_index {|elem, index| self.sheet.column_title_indexes[elem.text] = index+self.column_titles_start[1]-1}
    end

    def set_row_title_indexes
      self.sheet.row_title_indexes = {}
      self.row_titles.each_with_index {|elem, index| self.sheet.row_title_indexes[elem.text] = index+self.row_titles_start[0]-1}
    end

    def set_column_widths
      col_widths = {}

      element_strings = []

      self.elements.each do |element|
        longest_line = element.text.split("\n").max_by{|line| line.length}
        element_strings.push [element.col, longest_line || '']
        (1..element.merge).each {|column| col_widths[column] = 0} if element.merge
      end

      element_strings.each do |col, text|
        col_widths[col] ||= 0
        col_widths[col] = text.length if text.length > col_widths[col]
      end

      self.column_titles.each_with_index do |elem, index|
        col_widths[index+self.column_titles_start[1]-1] ||= 0
        col_widths[index+self.column_titles_start[1]-1] = elem.text.length if elem.text.length > col_widths[index+self.column_titles_start[1]-1]
      end

      unless self.row_titles.empty?
        col_widths[self.row_titles_start[1]-1] ||= 0
        max_row_title = self.row_titles.max_by{|elem| elem.text.length}.text.length
        col_widths[self.row_titles_start[1]-1] = max_row_title if max_row_title > col_widths[self.row_titles_start[1]-1]
      end

      self.column_data.each do |elem|
        title_column = self.sheet.column_title_indexes[elem.title]
        col_widths[title_column] ||= 0
        col_widths[title_column] = elem.text.length if elem.text.length > col_widths[title_column]
      end

      self.row_data.each do |elem|
        self.row_data.select{|data| data.text == elem.text}.each_with_index do |title_data, index|
          column = self.row_titles_start[1] + 1 + index
          col_widths[column] ||= 0
          col_widths[column] = title_data.text.length if title_data.text.length > col_widths[column]
        end
      end

      widths = (0..self.max[:col].to_i).map{|col| col_widths[col] ? col_widths[col]+4 : 0}
      self.sheet.column_widths *widths
    end

    def set_row_heights
      row_heights = {}

      self.elements.each do |elem|
        row_heights[elem.row] ||= 0
        height = elem.text.split("\n").count * 10 + 10
        row_heights[elem.row] = height if height > row_heights[elem.row]
      end

      unless self.column_titles.empty?
        max_column_title_height = self.column_titles.map{|elem| elem.text.split("\n").size}.max * 10 + 10
        row_heights[self.column_titles_start[0]] ||= 0
        row_heights[self.column_titles_start[0]] = max_column_title_height if max_column_title_height > row_heights[self.column_titles_start[0]]
      end


      (0..self.max[:row].to_i).each {|row| self.sheet.rows[row-1].height = row_heights[row] || 20}
    end

    def place_elements
      self.elements.each do |elem|
        self.sheet.name_to_cell("#{column(elem.col)}#{elem.row}").value = elem.text
        style = self.sheet.styles.add_style elem.combined_style
        self.sheet.name_to_cell("#{column(elem.col)}#{elem.row}").style = style
        self.sheet.add_comment ref: "#{column(elem.col)}#{elem.row}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
        self.sheet.merge_cells "#{column(elem.col)}#{elem.row}:#{column(elem.col+elem.merge)}#{elem.row}" if elem.merge
      end
    end

    def set_max_row_and_column
      self.max = {col: 0, row: 0}
      self.max[:row] = self.elements.max_by{|elem| elem.row}.row unless self.elements.empty?
      self.max[:col] = self.elements.max_by{|elem| elem.col}.col unless self.elements.empty?
      self.max[:col] = self.column_titles.count if self.column_titles and self.column_titles.count > self.max[:col]
      self.max[:row] = self.row_titles.count if self.row_titles and self.row_titles.count > self.max[:row]
      self.max[:row] += self.column_titles_start[0] if self.column_titles_start
      self.max[:col] += self.column_titles_start[1] if self.column_titles_start
      self.max[:row] += self.row_titles_start[0] if self.row_titles_start
      self.max[:col] += self.row_titles_start[1] if self.row_titles_start
      self.max[:row] += self.column_data.size if self.column_data
      self.max[:col] += self.row_data.size if self.row_data
      self.max[:col] += self.column_titles.select{|elem| elem.list}.size + self.row_titles.select{|elem| elem.list}.size
      (0..self.max[:row]+1).each {self.sheet.add_row(Array.new(self.max[:col].to_i, nil))}
    end

    def place_column_titles
      self.column_titles.each_with_index do |elem, index|
        cell = self.sheet.name_to_cell("#{column(index+self.column_titles_start[1]-1)}#{self.column_titles_start[0]}")
        cell.value = elem.text
        style = self.sheet.styles.add_style elem.combined_style
        cell.style = style
        self.sheet.add_hyperlink location: elem.hyperlink, ref: cell if elem.hyperlink
        self.sheet.add_comment ref: "#{column(index+self.column_titles_start[1]-1)}#{self.column_titles_start[0]}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
      end
    end

    def set_lists_for_column_titles
      existing_titles_with_lists = self.column_titles.select{|elem| elem.list}
      return false if existing_titles_with_lists.empty?
      if lists_sheet = current_lists_sheet
        pre_existing_titles_with_lists = []
        lists_sheet.column_title_indexes.each do |title, column|
          row = 2
          list = []
          while cell = lists_sheet.name_to_cell("#{column column}#{row}") and cell.value
            list << cell.value
            row += 1
          end
          pre_existing_titles_with_lists << SheetBuilder::Title.new(text: title, list: list)
        end
        titles_with_lists = pre_existing_titles_with_lists + existing_titles_with_lists
        lists_sheet_index = self.sheet.workbook.worksheets.index {|sheet| sheet.name == 'Lists'}
        self.sheet.workbook.worksheets.delete_at lists_sheet_index
      else
        titles_with_lists = existing_titles_with_lists
      end
      lists_sheet = self.sheet.workbook.add_worksheet name: 'Lists'
      list_titles = []
      list_data = []
      titles_with_lists.each do |elem|
        list_titles = titles_with_lists.map{|elem| {text: elem.text}}
        elem.list.each {|list_item| list_data << {text: list_item, title: elem.text}}
      end
      list_titles.uniq!
      list_data.uniq!
      blueprint = SheetBuilder::Blueprint.new column_titles: list_titles
      blueprint.build! sheet: lists_sheet, column_data: list_data
      existing_titles_with_lists.each do |elem|
        100.times do |row|
          list_column = current_lists_sheet.column_title_indexes[elem.text]
          self.sheet.add_data_validation("#{column (self.sheet.column_title_indexes[elem.text])}#{self.column_titles_start[0]+row+1}", {
              type: :list,
              formula1: "Lists!#{column list_column}2:#{column list_column}#{elem.list.size+1}",
              showDropDown: false,
              showErrorMessage: true,
              errorTitle: '',
              errorStyle: :stop,
              showInputMessage: true})
        end
      end
    end

    def place_row_titles
      self.row_titles.each_with_index do |elem, index|
        cell = self.sheet.name_to_cell("#{column(self.row_titles_start[1]-1)}#{self.row_titles_start[0]+index}")
        cell.value = elem.text
        style = self.sheet.styles.add_style elem.combined_style
        cell.style = style
        self.sheet.add_hyperlink location: elem.hyperlink, ref: cell if elem.hyperlink
        self.sheet.add_comment ref: "#{column(self.row_titles_start[1]-1)}#{self.row_titles_start[0]+index}", text: "#{elem.comment}", author: elem.text, visible: false if elem.comment
      end
    end

    def set_lists_for_row_titles
      existing_titles_with_lists = self.row_titles.select{|elem| elem.list}
      return false if existing_titles_with_lists.empty?
      if lists_sheet = current_lists_sheet
        pre_existing_titles_with_lists = []
        lists_sheet.column_title_indexes.each do |title, column|
          row = 2
          list = []
          while cell = lists_sheet.name_to_cell("#{column column}#{row}") and cell.value
            list << cell.value
            row += 1
          end
          pre_existing_titles_with_lists << SheetBuilder::Title.new(text: title, list: list)
        end
        titles_with_lists = pre_existing_titles_with_lists + existing_titles_with_lists
        lists_sheet_index = self.sheet.workbook.worksheets.index {|sheet| sheet.name == 'Lists'}
        self.sheet.workbook.worksheets.delete_at lists_sheet_index
      else
        titles_with_lists = existing_titles_with_lists
      end
      lists_sheet = self.sheet.workbook.add_worksheet name: 'Lists'
      list_titles = []
      list_data = []
      titles_with_lists.each do |elem|
        list_titles = titles_with_lists.map{|elem| {text: elem.text}}
        elem.list.each {|list_item| list_data << {text: list_item, title: elem.text}}
      end
      list_titles.uniq!
      list_data.uniq!
      blueprint = SheetBuilder::Blueprint.new column_titles: list_titles
      blueprint.build! sheet: lists_sheet, column_data: list_data
      existing_titles_with_lists.each do |elem|
        100.times do |column|
          list_column = current_lists_sheet.column_title_indexes[elem.text]
          self.sheet.add_data_validation("#{column self.row_titles_start[1]+column}#{self.sheet.row_title_indexes[elem.text]+1}", {
              type: :list,
              formula1: "Lists!#{column list_column}2:#{column list_column}#{elem.list.size+1}",
              showDropDown: false,
              showErrorMessage: true,
              errorTitle: '',
              errorStyle: :stop,
              showInputMessage: true})
        end
      end
    end


    def set_column_title_row_height
      return false if self.column_titles.empty?
      if self.column_title_row_height
        self.sheet.rows[self.column_titles_start[0]-1].height = column_title_row_height
      else
        most_lines = self.column_titles.max_by{|elem| elem.text.split("\n").size}.text.split("\n").size
        self.sheet.rows[self.column_titles_start[0]-1].height = most_lines * 10 + 10
      end
    end

    def place_column_data
      self.column_data.each do |elem|
        title_column = column self.sheet.column_title_indexes[elem.title]
        if elem.row
          self.sheet.name_to_cell("#{title_column}#{elem.row + self.column_titles_start[0] + 1}").value = elem.text
        else
          row = self.column_titles_start[0] + 1
          while self.sheet.name_to_cell("#{title_column}#{row}").value
            row += 1
          end
          self.sheet.name_to_cell("#{title_column}#{row}").value = elem.text
        end
      end
    end

    def place_row_data
      self.row_data.each do |elem|
        title_row = self.sheet.row_title_indexes[elem.title]
        column = self.row_titles_start[1]
        while self.sheet.name_to_cell("#{column column}#{title_row+1}").value
          column += 1
        end
        self.sheet.name_to_cell("#{column column}#{title_row+1}").value = elem.text
      end
    end

    def move_lists_sheet_to_end
      return false if self.sheet.name == 'Lists'
      sheets = self.sheet.workbook.worksheets
      lists_sheet_index = sheets.index { |sheet| sheet.name == 'Lists' }
      if lists_sheet_index
        lists_sheet = sheets[lists_sheet_index]
        sheets.delete_at lists_sheet_index
        sheets << lists_sheet
        sheets.each_with_index {|sheet, index| sheet.workbook.worksheets[index] = sheet}
      end
    end

    def column index
      column = (index % 26 + 65).chr
      column << (index / 26 + 64).chr if index > 25
      column.reverse
    end

    def current_lists_sheet
      self.sheet.workbook.sheet_by_name 'Lists'
    end

  end

end

class Axlsx::Worksheet
  attr_accessor :column_title_indexes
  attr_accessor :row_title_indexes
end
