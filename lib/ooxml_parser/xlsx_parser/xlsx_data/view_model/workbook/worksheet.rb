require_relative 'worksheet/excel_comments'
require_relative 'worksheet/ole_objects'
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
    # @return [Relationships] array of ole objects
    attr_accessor :ole_objects
    # @return [ExtensionList] list of extensions
    attr_accessor :extension_list

    def initialize(parent: nil)
      @columns = []
      @name = ''
      @rows = []
      @merge = []
      @charts = []
      @hyperlinks = []
      @drawings = []
      @sheet_views = []
      @table_parts = []
      @parent = parent
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

    def parse(path_to_xml_file)
      @xml_name = File.basename path_to_xml_file
      parse_relationships
      OOXMLDocumentObject.add_to_xmls_stack("#{OOXMLDocumentObject.root_subfolder}/worksheets/#{File.basename(path_to_xml_file)}")
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      sheet = doc.search('//xmlns:worksheet').first
      sheet.xpath('*').each do |worksheet_node_child|
        case worksheet_node_child.name
        when 'sheetData'
          worksheet_node_child.xpath('xmlns:row').each do |row_node|
            @rows[row_node.attribute('r').value.to_i - 1] = XlsxRow.new(parent: self).parse(row_node)
            @rows[row_node.attribute('r').value.to_i - 1].style = root_object.style_sheet.cell_xfs.xf_array[row_node.attribute('s').value.to_i].calculate_values unless row_node.attribute('s').nil?
          end
        when 'sheetFormatPr'
          if !worksheet_node_child.attribute('defaultColWidth').nil? && !worksheet_node_child.attribute('defaultRowHeight').nil?
            @sheet_format_properties = SheetFormatProperties.new(parent: self).parse(worksheet_node_child)
          end
        when 'mergeCells'
          worksheet_node_child.xpath('xmlns:mergeCell').each do |merge_node|
            @merge << merge_node.attribute('ref').value.to_s
          end
        when 'drawing'
          path_to_drawing = OOXMLDocumentObject.get_link_from_rels(worksheet_node_child.attribute('id').value)
          unless path_to_drawing.nil?
            OOXMLDocumentObject.add_to_xmls_stack(path_to_drawing)
            parse_drawing
            OOXMLDocumentObject.xmls_stack.pop
          end
        when 'hyperlinks'
          worksheet_node_child.xpath('xmlns:hyperlink').each do |hyperlink_node|
            @hyperlinks << Hyperlink.new(parent: self).parse(hyperlink_node).dup
          end
        when 'cols'
          @columns = XlsxColumnProperties.parse_list(worksheet_node_child, parent: self)
        when 'autoFilter'
          @autofilter = Autofilter.new(parent: self).parse(worksheet_node_child)
        when 'tableParts'
          worksheet_node_child.xpath('*').each do |part_node|
            @table_parts << TablePart.new(parent: self).parse(part_node)
          end
        when 'sheetViews'
          worksheet_node_child.xpath('*').each do |view_child|
            @sheet_views << SheetView.new(parent: self).parse(view_child)
          end
        when 'oleObjects'
          @ole_objects = OleObjects.new(parent: self).parse(worksheet_node_child)
        when 'extLst'
          @extension_list = ExtensionList.new(parent: self).parse(worksheet_node_child)
        end
      end
      @comments = ExcelComments.parse_file(File.basename(path_to_xml_file), OOXMLDocumentObject.path_to_folder)
      OOXMLDocumentObject.xmls_stack.pop
      self
    end
  end
end
