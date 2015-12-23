require_relative 'worksheet/excel_comments'
require_relative 'worksheet/sheet_format_properties'
require_relative 'worksheet/sheet_view'
require_relative 'worksheet/xlsx_column_properties'
require_relative 'worksheet/xlsx_drawing'
require_relative 'worksheet/xlsx_row'
require_relative 'worksheet/xlsx_table'
# Properties of worksheet
module OoxmlParser
  class Worksheet < OOXMLDocumentObject
    attr_accessor :name, :rows, :merge, :charts, :hyperlinks, :drawings, :comments, :columns, :sheet_format_properties,
                  :autofilter, :table_parts, :sheet_views

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

    def self.parse(path_to_xml_file)
      worksheet = Worksheet.new
      @current_sheet_xml_name = File.basename path_to_xml_file
      OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/worksheets/#{File.basename(path_to_xml_file)}")
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      sheet = doc.search('//xmlns:worksheet').first
      sheet.xpath('*').each do |worksheet_node_child|
        case worksheet_node_child.name
        when 'sheetData'
          worksheet_node_child.xpath('xmlns:row').each do |row_node|
            worksheet.rows[row_node.attribute('r').value.to_i - 1] = XlsxRow.parse(row_node)
            worksheet.rows[row_node.attribute('r').value.to_i - 1].style = CellStyle.parse(row_node.attribute('s').value) unless row_node.attribute('s').nil?
          end
        when 'sheetFormatPr'
          if !worksheet_node_child.attribute('defaultColWidth').nil? && !worksheet_node_child.attribute('defaultRowHeight').nil?
            worksheet.sheet_format_properties = SheetFormatProperties.parse(worksheet_node_child)
          end
        when 'mergeCells'
          worksheet_node_child.xpath('xmlns:mergeCell').each do |merge_node|
            worksheet.merge << merge_node.attribute('ref').value.to_s
          end
        when 'drawing'
          path_to_drawing = OOXMLDocumentObject.get_link_from_rels(worksheet_node_child.attribute('id').value)
          unless path_to_drawing.nil?
            OOXMLDocumentObject.add_to_xmls_stack(path_to_drawing)
            XlsxDrawing.parse_list(worksheet)
            OOXMLDocumentObject.xmls_stack.pop
          end
        when 'hyperlinks'
          worksheet_node_child.xpath('xmlns:hyperlink').each do |hyperlink_node|
            worksheet.hyperlinks << Hyperlink.parse(hyperlink_node).dup
          end
        when 'cols'
          worksheet.columns = XlsxColumnProperties.parse_list(worksheet_node_child)
        when 'autoFilter'
          worksheet.autofilter = Coordinates.parser_coordinates_range(worksheet_node_child.attribute('ref').value.to_s)
        when 'tableParts'
          worksheet.table_parts = []
          worksheet_node_child.xpath('xmlns:tablePart').each { |table_part_node| worksheet.table_parts << XlsxTable.parse(table_part_node) }
        when 'sheetViews'
          worksheet_node_child.xpath('xmlns:sheetView').each { |sheet_view_node| worksheet.sheet_views << SheetView.parse(sheet_view_node) }
        end
      end
      worksheet.comments = ExcelComments.parse_file(File.basename(path_to_xml_file), OOXMLDocumentObject.path_to_folder)
      OOXMLDocumentObject.xmls_stack.pop
      worksheet
    end
  end
end
