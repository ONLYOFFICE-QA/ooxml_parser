require_relative 'worksheet/excel_comments'
require_relative 'worksheet/sheet_format_properties'
require_relative 'worksheet/sheet_view'
require_relative 'worksheet/table_part'
require_relative 'worksheet/xlsx_column_properties'
require_relative 'worksheet/xlsx_drawing'
require_relative 'worksheet/xlsx_row'
# Properties of worksheet
module OoxmlParser
  class Worksheet < OOXMLDocumentObject
    attr_accessor :name, :rows, :merge, :charts, :hyperlinks, :drawings, :comments, :columns, :sheet_format_properties,
                  :autofilter, :table_parts, :sheet_views
    # @return [String] xml name of sheet
    attr_accessor :xml_name
    # @return [Relationships] array of relationships
    attr_accessor :relationships

    def initialize
      @columns = []
      @name = ''
      @rows = []
      @merge = []
      @charts = []
      @hyperlinks = []
      @drawings = []
      @sheet_views = []
      @table_parts = []
    end

    def parse_relationships
      OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/worksheets/_rels/#{@xml_name}.rels")
      @relationships = Relationships.parse_rels(OOXMLDocumentObject.current_xml) if File.exist?(OOXMLDocumentObject.current_xml)
      OOXMLDocumentObject.xmls_stack.pop
    end

    # @return [True, false] if structure contain any user data
    def with_data?
      return true unless @rows.empty?
      return true unless @columns.empty?
      return true unless @drawings.empty?
      return true unless @charts.empty?
      return true unless @hyperlinks.empty?
      false
    end

    # Parse list of drawings in file
    def parse_drawing
      drawing_node = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      drawing_node.xpath('xdr:wsDr/*').each do |drawing_node_child|
        @drawings << XlsxDrawing.new(parent: self).parse(drawing_node_child)
      end
    end

    def self.parse(path_to_xml_file, parent: nil)
      worksheet = Worksheet.new
      worksheet.xml_name = File.basename path_to_xml_file
      worksheet.parse_relationships
      worksheet.parent = parent
      OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/worksheets/#{File.basename(path_to_xml_file)}")
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      sheet = doc.search('//xmlns:worksheet').first
      sheet.xpath('*').each do |worksheet_node_child|
        case worksheet_node_child.name
        when 'sheetData'
          worksheet_node_child.xpath('xmlns:row').each do |row_node|
            worksheet.rows[row_node.attribute('r').value.to_i - 1] = XlsxRow.new(parent: worksheet).parse(row_node)
            worksheet.rows[row_node.attribute('r').value.to_i - 1].style = CellStyle.new(parent: worksheet).parse(row_node.attribute('s').value) unless row_node.attribute('s').nil?
          end
        when 'sheetFormatPr'
          if !worksheet_node_child.attribute('defaultColWidth').nil? && !worksheet_node_child.attribute('defaultRowHeight').nil?
            worksheet.sheet_format_properties = SheetFormatProperties.new(parent: worksheet).parse(worksheet_node_child)
          end
        when 'mergeCells'
          worksheet_node_child.xpath('xmlns:mergeCell').each do |merge_node|
            worksheet.merge << merge_node.attribute('ref').value.to_s
          end
        when 'drawing'
          path_to_drawing = OOXMLDocumentObject.get_link_from_rels(worksheet_node_child.attribute('id').value)
          unless path_to_drawing.nil?
            OOXMLDocumentObject.add_to_xmls_stack(path_to_drawing)
            worksheet.parse_drawing
            OOXMLDocumentObject.xmls_stack.pop
          end
        when 'hyperlinks'
          worksheet_node_child.xpath('xmlns:hyperlink').each do |hyperlink_node|
            worksheet.hyperlinks << Hyperlink.new(parent: worksheet).parse(hyperlink_node).dup
          end
        when 'cols'
          worksheet.columns = XlsxColumnProperties.parse_list(worksheet_node_child, parent: worksheet)
        when 'autoFilter'
          worksheet.autofilter = Coordinates.parser_coordinates_range(worksheet_node_child.attribute('ref').value.to_s)
        when 'tableParts'
          worksheet_node_child.xpath('*').each do |part_node|
            worksheet.table_parts << TablePart.new(parent: worksheet).parse(part_node)
          end
        when 'sheetViews'
          worksheet_node_child.xpath('*').each do |view_child|
            worksheet.sheet_views << SheetView.new(parent: worksheet).parse(view_child)
          end
        end
      end
      worksheet.comments = ExcelComments.parse_file(File.basename(path_to_xml_file), OOXMLDocumentObject.path_to_folder)
      OOXMLDocumentObject.xmls_stack.pop
      worksheet
    end
  end
end
