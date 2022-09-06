# frozen_string_literal: true

require_relative 'workbook/chartsheet'
require_relative 'workbook/pivot_cache'
require_relative 'workbook/pivot_table_definition'
require_relative 'workbook/defined_name'
require_relative 'workbook/workbook_protection'
require_relative 'workbook/shared_string_table'
require_relative 'workbook/style_sheet'
require_relative 'workbook/worksheet'
require_relative 'workbook/sheet'
require_relative 'workbook/workbook_helpers'
module OoxmlParser
  # Class for storing XLSX Workbook
  class XLSXWorkbook < CommonDocumentStructure
    include WorkbookHelpers
    attr_accessor :worksheets
    # @return [Array<Sheet>] list of sheets
    attr_reader :sheets
    # @return [PresentationTheme] theme of Workbook
    attr_accessor :theme
    # @return [Relationships] rels of book
    attr_accessor :relationships
    # @return [StyleSheet] styles of book
    attr_accessor :style_sheet
    # @return [SharedStringTable] styles of book
    attr_accessor :shared_strings_table
    # @return [Array<PivotCache>] list of pivot caches
    attr_accessor :pivot_caches
    # @return [Array<PivotTableDefintion>] list of pivot table defitions
    attr_accessor :pivot_table_definitions
    # @return [Array<DefinedName>] list of defined names
    attr_reader :defined_names
    # @return [WorkbookProtection] protection of workbook structure
    attr_reader :workbook_protection

    def initialize(params = {})
      @worksheets = []
      @sheets = []
      @pivot_caches = []
      @pivot_table_definitions = []
      @defined_names = []
      super
    end

    # Return cell by coordinates
    # @param column [String, Integer] number or numeric digit of column
    # @param row [Integer] row to find
    # @param sheet [Integer] number of sheet
    # @return [XlsxCell] result
    def cell(column, row, sheet = 0)
      column = Coordinates.new(row, column).column_number unless StringHelper.numeric?(column.to_s)

      if StringHelper.numeric?(sheet.to_s)
        row = @worksheets[sheet].rows[row.to_i - 1]
        return nil if row.nil?

        return row.cells[column.to_i - 1]
      elsif sheet.is_a?(String)
        @worksheets.each do |worksheet|
          next unless worksheet.name == sheet
          next unless worksheet.rows[row.to_i - 1]

          return worksheet.rows[row.to_i - 1].cells[column.to_i - 1]
        end
        return nil
      end
      raise "Error. Wrong sheet value: #{sheet}"
    end

    # Get all values of formulas.
    # @param [Fixnum] precision of formulas counting
    # @return [Array, String] all formulas
    def all_formula_values(precision = 14)
      formulas = []
      worksheets.each do |c_sheet|
        next unless c_sheet

        c_sheet.rows.each do |c_row|
          next unless c_row

          c_row.cells.each do |c_cell|
            next unless c_cell
            next unless c_cell.formula
            next if c_cell.formula.empty?

            text = c_cell.raw_text
            if StringHelper.numeric?(text)
              text = text.to_f.round(10).to_s[0..precision]
            elsif StringHelper.complex?(text)
              complex_number = Complex(text.tr(',', '.'))
              real_part = complex_number.real
              real_rounded = real_part.to_f.round(10).to_s[0..precision].to_f

              imag_part = complex_number.imag
              imag_rounded = imag_part.to_f.round(10).to_s[0..precision].to_f
              complex_rounded = Complex(real_rounded, imag_rounded)
              text = complex_rounded.to_s
            end
            formulas << text
          end
        end
      end
      formulas
    end

    # Do work for parsing shared strings file
    def parse_shared_strings
      shared_strings_target = relationships.target_by_type('sharedString')
      return if shared_strings_target.empty?

      shared_string_file = "#{root_object.unpacked_folder}/xl/#{shared_strings_target.first}"
      @shared_strings_table = SharedStringTable.new(parent: self).parse(shared_string_file)
    end

    # Parse content of Workbook
    # @return [XLSXWorkbook]
    def parse
      @content_types = ContentTypes.new(parent: self).parse
      @relationships = Relationships.new(parent: self).parse_file("#{root_object.unpacked_folder}xl/_rels/workbook.xml.rels")
      parse_shared_strings
      @root_subfolder = 'xl/'
      root_object.add_to_xmls_stack('xl/workbook.xml')
      @doc = Nokogiri::XML.parse(File.open(root_object.current_xml))
      @theme = PresentationTheme.new(parent: self).parse("xl/#{link_to_theme_xml}") if link_to_theme_xml
      @style_sheet = StyleSheet.new(parent: self).parse
      @doc.xpath('xmlns:workbook/xmlns:sheets/xmlns:sheet').each do |sheet|
        @sheets << Sheet.new(parent: self).parse(sheet)
        file = @relationships.target_by_id(sheet.attribute('id').value)
        if file.start_with?('worksheets')
          @worksheets << Worksheet.new(parent: self).parse(file)
          @worksheets.last.name = @sheets.last.name
        elsif file.start_with?('chartsheets')
          @worksheets << Chartsheet.new(parent: self).parse(file)
        end
      end
      parse_pivot_cache
      parse_pivot_table
      parse_defined_names
      parse_workbook_protection
      root_object.xmls_stack.pop
      self
    end

    private

    def link_to_theme_xml
      relationships.target_by_type('theme').first
    end

    # Perform parsing of pivot cache
    def parse_pivot_cache
      @doc.xpath('xmlns:workbook/xmlns:pivotCaches/xmlns:pivotCache').each do |pivot_cache|
        @pivot_caches << PivotCache.new(parent: self).parse(pivot_cache)
      end
    end

    # Perform parsing of pivot table
    def parse_pivot_table
      files = @content_types.by_type('application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml')
      files.each do |file|
        @pivot_table_definitions << PivotTableDefinition.new(parent: self).parse(file.part_name)
      end
    end

    # Perform parsing of defined names
    def parse_defined_names
      @doc.xpath('xmlns:workbook/xmlns:definedNames/xmlns:definedName').each do |defined_name|
        @defined_names << DefinedName.new(parent: self).parse(defined_name)
      end
    end

    # Perform parsing of workbook protection
    def parse_workbook_protection
      workbook_protection = @doc.search('//xmlns:workbookProtection').first
      return nil unless workbook_protection

      @workbook_protection = WorkbookProtection.new(parent: self).parse(workbook_protection)
    end
  end
end
