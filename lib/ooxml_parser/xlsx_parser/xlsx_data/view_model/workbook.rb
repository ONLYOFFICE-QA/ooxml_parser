require_relative 'workbook/worksheet'
# Class for storing XLSX Workbook
module OoxmlParser
  class XLSXWorkbook < OOXMLDocumentObject
    attr_accessor :worksheets

    def initialize(worksheets = [])
      @worksheets = worksheets
    end

    def cell(column, row, sheet = 0)
      column = Coordinates.new(row, column).get_column_number unless StringHelper.numeric?(column.to_s)

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

    def self.parse(path_to_folder)
      workbook = XLSXWorkbook.new
      OOXMLDocumentObject.path_to_folder = path_to_folder
      OOXMLDocumentObject.xmls_stack = []
      OOXMLDocumentObject.root_subfolder = 'xl/'
      self.shared_strings = nil
      OOXMLDocumentObject.add_to_xmls_stack('xl/workbook.xml')
      doc = Nokogiri::XML.parse(File.open(OOXMLDocumentObject.current_xml))
      XLSXWorkbook.styles_node = Nokogiri::XML(File.open("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/styles.xml"))
      PresentationTheme.parse("xl/#{link_to_theme_xml}") if link_to_theme_xml
      doc.xpath('xmlns:workbook/xmlns:sheets/xmlns:sheet').each do |sheet|
        workbook.worksheets << Worksheet.parse('xl/' + OOXMLDocumentObject.get_link_from_rels(sheet.attribute('id').value))
        workbook.worksheets.last.name = sheet.attribute('name').value
      end
      OOXMLDocumentObject.xmls_stack.pop
      workbook
    end

    class << self
      # @return [Array, Nokogiri::XML::Eelement] list of shared strings
      attr_accessor :shared_strings
      attr_accessor :styles_node

      # Accessor for shared string. Initialization for this array
      def shared_strings
        if @shared_strings.nil?
          shared_strings_file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/sharedStrings.xml"
          @shared_strings = Nokogiri::XML(File.open(shared_strings_file)).xpath('//xmlns:si') if File.exist?(shared_strings_file)
        end
        @shared_strings
      end

      def link_to_theme_xml
        doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'xl/_rels/workbook.xml.rels'))
        doc.xpath('xmlns:Relationships/xmlns:Relationship').each do |relationship_node|
          if relationship_node.attribute('Target').value.include?('theme')
            return relationship_node.attribute('Target').value
          end
        end
        nil
      end
    end
  end
end
