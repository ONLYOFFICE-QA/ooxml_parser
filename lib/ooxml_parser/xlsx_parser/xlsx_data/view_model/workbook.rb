# frozen_string_literal: true

require_relative 'workbook/chartsheet'
require_relative 'workbook/shared_string_table'
require_relative 'workbook/style_sheet'
require_relative 'workbook/worksheet'
require_relative 'workbook/workbook_helpers'
# Class for storing XLSX Workbook
module OoxmlParser
  class XLSXWorkbook < CommonDocumentStructure
    include WorkbookHelpers
    attr_accessor :worksheets
    # @return [PresentationTheme] theme of Workbook
    attr_accessor :theme
    # @return [Relationships] rels of book
    attr_accessor :relationships
    # @return [StyleSheet] styles of book
    attr_accessor :style_sheet
    # @return [SharedStringTable] styles of book
    attr_accessor :shared_strings_table

    def initialize(params = {})
      @worksheets = []
      super
    end

    def cell(column, row, sheet = 0)
      column = Coordinates.new(row, column).column_number unless StringHelper.numeric?(column.to_s)

      if StringHelper.numeric?(sheet.to_s)
        row = @worksheets[sheet].rows[row.to_i - 1]
        return nil if row.nil?

        return row.cells[column.to_i - 1]
      elsif sheet.is_a?(String)
        @worksheets.each do |worksheet|
          if worksheet.name == sheet
            return worksheet.rows[row.to_i - 1].cells[column.to_i - 1] unless worksheet.rows[row.to_i - 1].nil?
          end
        end
        return nil
      end
      raise "Error. Wrong sheet value: #{sheet}"
    end

    def difference(other)
      Hash.object_to_hash(self).diff(Hash.object_to_hash(other))
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

      shared_string_file = "#{OOXMLDocumentObject.path_to_folder}/xl/#{shared_strings_target.first}"
      @shared_strings_table = SharedStringTable.new(parent: self).parse(shared_string_file)
    end

    def parse
      @content_types = ContentTypes.new(parent: self).parse
      @relationships = Relationships.new(parent: self).parse_file("#{OOXMLDocumentObject.path_to_folder}xl/_rels/workbook.xml.rels")
      parse_shared_strings
      OOXMLDocumentObject.xmls_stack = []
      OOXMLDocumentObject.root_subfolder = 'xl/'
      OOXMLDocumentObject.add_to_xmls_stack('xl/workbook.xml')
      doc = Nokogiri::XML.parse(File.open(OOXMLDocumentObject.current_xml))
      @theme = PresentationTheme.parse("xl/#{link_to_theme_xml}") if link_to_theme_xml
      @style_sheet = StyleSheet.new(parent: self).parse
      doc.xpath('xmlns:workbook/xmlns:sheets/xmlns:sheet').each do |sheet|
        file = @relationships.target_by_id(sheet.attribute('id').value)
        if file.start_with?('worksheets')
          @worksheets << Worksheet.new(parent: self).parse(file)
          @worksheets.last.name = sheet.attribute('name').value
        elsif file.start_with?('chartsheets')
          @worksheets << Chartsheet.new(parent: self).parse(file)
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      self
    end

    private

    def link_to_theme_xml
      relationships.target_by_type('theme').first
    end
  end
end
