# frozen_string_literal: true

require_relative 'worksheet/excel_comments'
require_relative 'worksheet/ole_objects'
require_relative 'worksheet/page_setup'
require_relative 'worksheet/sheet_format_properties'
require_relative 'worksheet/sheet_view'
require_relative 'worksheet/table_part'
require_relative 'worksheet/worksheet_helper'
require_relative 'worksheet/xlsx_column_properties'
require_relative 'worksheet/xlsx_drawing'
require_relative 'worksheet/xlsx_row'
require_relative 'worksheet/xlsx_header_footer'
require_relative 'worksheet/sheet_protection'
require_relative 'worksheet/protected_range'
module OoxmlParser
  # Properties of worksheet
  class Worksheet < OOXMLDocumentObject
    include WorksheetHelper
    attr_accessor :name, :merge, :charts, :hyperlinks, :drawings, :comments, :columns, :sheet_format_properties,
                  :autofilter, :table_parts, :sheet_views
    # @return [String] xml name of sheet
    attr_accessor :xml_name
    # @return [Relationships] array of relationships
    attr_accessor :relationships
    # @return [Relationships] array of ole objects
    attr_accessor :ole_objects
    # @return [PageMargins] page margins settings
    attr_reader :page_margins
    # @return [PageSetup] page setup settings
    attr_reader :page_setup
    # @return [ExtensionList] list of extensions
    attr_accessor :extension_list
    # @return [XlsxHeaderFooter] header and footer
    attr_reader :header_footer
    # @return [Array<ConditionalFormatting>] list of conditional formattings
    attr_reader :conditional_formattings
    # @return [SheetProtection] protection of sheet
    attr_reader :sheet_protection
    # @return [Array<ProtectedRange>] list of protected ranges
    attr_reader :protected_ranges
    # @return [Array<Row>] rows in sheet, as in xml structure
    attr_reader :rows_raw

    def initialize(parent: nil)
      @columns = []
      @name = ''
      @rows = []
      @rows_raw = []
      @merge = []
      @charts = []
      @hyperlinks = []
      @drawings = []
      @sheet_views = []
      @table_parts = []
      @conditional_formattings = []
      @protected_ranges = []
      super
    end

    # Perform parsing of relationships
    # @return [nil]
    def parse_relationships
      root_object.add_to_xmls_stack("#{root_object.root_subfolder}/worksheets/_rels/#{@xml_name}.rels")
      @relationships = Relationships.new(parent: self).parse_file(root_object.current_xml) if File.exist?(root_object.current_xml)
      root_object.xmls_stack.pop
    end

    # @return [True, false] if structure contain any user data
    def with_data?
      return true unless @rows_raw.empty?
      return true unless default_columns?
      return true unless @drawings.empty?
      return true unless @charts.empty?
      return true unless @hyperlinks.empty?

      false
    end

    # Parse list of drawings in file
    def parse_drawing
      drawing_node = parse_xml(root_object.current_xml)
      drawing_node.xpath('xdr:wsDr/*').each do |drawing_node_child|
        @drawings << XlsxDrawing.new(parent: self).parse(drawing_node_child)
      end
    end

    # Parse data of Worksheet
    # @param path_to_xml_file [String] path to file to parse
    # @return [Worksheet] parsed worksheet
    def parse(path_to_xml_file)
      @xml_name = File.basename path_to_xml_file
      parse_relationships
      root_object.add_to_xmls_stack("#{root_object.root_subfolder}/worksheets/#{File.basename(path_to_xml_file)}")
      doc = parse_xml(root_object.current_xml)
      sheet = doc.search('//xmlns:worksheet').first
      sheet.xpath('*').each do |worksheet_node_child|
        case worksheet_node_child.name
        when 'sheetData'
          worksheet_node_child.xpath('xmlns:row').each do |row_node|
            @rows_raw << XlsxRow.new(parent: self).parse(row_node)
          end
        when 'sheetFormatPr'
          @sheet_format_properties = SheetFormatProperties.new(parent: self).parse(worksheet_node_child)
        when 'mergeCells'
          worksheet_node_child.xpath('xmlns:mergeCell').each do |merge_node|
            @merge << merge_node.attribute('ref').value.to_s
          end
        when 'drawing'
          @drawing = DocxDrawing.new(parent: self).parse(worksheet_node_child)
          path_to_drawing = root_object.get_link_from_rels(@drawing.id)
          unless path_to_drawing.nil?
            root_object.add_to_xmls_stack(path_to_drawing)
            parse_drawing
            root_object.xmls_stack.pop
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
        when 'pageMargins'
          @page_margins = PageMargins.new(parent: self).parse(worksheet_node_child, :inch)
        when 'pageSetup'
          @page_setup = PageSetup.new(parent: self).parse(worksheet_node_child)
        when 'extLst'
          @extension_list = ExtensionList.new(parent: self).parse(worksheet_node_child)
        when 'headerFooter'
          @header_footer = XlsxHeaderFooter.new(parent: self).parse(worksheet_node_child)
        when 'conditionalFormatting'
          @conditional_formattings << ConditionalFormatting.new(parent: self).parse(worksheet_node_child)
        when 'sheetProtection'
          @sheet_protection = SheetProtection.new(parent: self).parse(worksheet_node_child)
        when 'protectedRanges'
          worksheet_node_child.xpath('*').each do |protected_range_node|
            @protected_ranges << ProtectedRange.new(parent: self).parse(protected_range_node)
          end
        end
      end
      parse_comments
      root_object.xmls_stack.pop
      self
    end

    # @return [Array<XlsxRow, nil>] list of rows, with nil,
    #   if row data is not stored in xml
    def rows
      return @rows if @rows.any?

      rows_raw.each do |row|
        @rows[row.index - 1] = row
      end

      @rows
    end

    private

    # Do work for parsing shared comments file
    def parse_comments
      return unless relationships

      comments_target = relationships.target_by_type('comment')
      return if comments_target.empty?

      comments_file = "#{root_object.unpacked_folder}/#{root_object.root_subfolder}/#{comments_target.first.gsub('..', '')}"
      @comments = ExcelComments.new(parent: self).parse(comments_file)
    end
  end
end
