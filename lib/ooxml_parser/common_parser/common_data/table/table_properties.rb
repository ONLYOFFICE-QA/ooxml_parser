# Table Properties Data
require_relative 'properties/table_layout'
require_relative 'properties/table_look'
require_relative 'properties/table_style'
require_relative 'properties/table_style_column_band_size'
require_relative 'properties/table_style_row_band_size'
require_relative 'properties/table_position'
require_relative 'properties/table_style_properties'
require_relative 'table_properties/table_borders'
module OoxmlParser
  # Class for parsing `w:tblPr` tags
  class TableProperties < OOXMLDocumentObject
    attr_accessor :jc, :table_width, :table_borders, :table_positon, :table_cell_margin, :table_indent, :stretching, :table_style, :row_banding_size,
                  :column_banding_size, :table_look, :grid_column, :style
    # @return [TableStyleColumnBandSize] table style column band size
    attr_accessor :table_style_column_band_size
    # @return [TableStyleRowBandSize] table style row band size
    attr_accessor :table_style_row_band_size
    # @return [TableLayout] table layout
    attr_accessor :table_layout
    # @return [OoxmlSize] table cell spacing
    attr_accessor :table_cell_spacing
    # @return [Shade] shade color of table
    attr_accessor :shade
    # @return [ValuedChild] caption of table
    attr_accessor :caption
    # @return [ValuedChild] description of table
    attr_accessor :description

    alias table_properties table_positon

    def initialize(parent: nil)
      @jc = :left
      @table_borders = TableBorders.new
      @parent = parent
    end

    # Parse TableProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tableStyleId'
          @style = TableStyle.new(parent: self).parse(style_id: node_child.text)
        when 'tblBorders'
          @table_borders = TableBorders.new(parent: self).parse(node_child)
        when 'tblStyle'
          @table_style = root_object.document_style_by_id(node_child.attribute('val').value)
        when 'tblW'
          @table_width = OoxmlSize.new.parse(node_child)
        when 'jc'
          @jc = node_child.attribute('val').text.to_sym
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'tblLook'
          @table_look = TableLook.new(parent: self).parse(node_child)
        when 'tblInd'
          @table_indent = OoxmlSize.new(node_child.attribute('w').text.to_f)
        when 'tblpPr'
          @table_positon = TablePosition.new(parent: self).parse(node_child)
        when 'tblCellMar'
          @table_cell_margin = TableMargins.new(parent: table_properties).parse(node_child)
        when 'tblStyleColBandSize'
          @table_style_column_band_size = TableStyleColumnBandSize.new(parent: self).parse(node_child)
        when 'tblStyleRowBandSize'
          @table_style_row_band_size = TableStyleRowBandSize.new(parent: self).parse(node_child)
        when 'tblLayout'
          @table_layout = TableLayout.new(parent: self).parse(node_child)
        when 'tblCellSpacing'
          @table_cell_spacing = OoxmlSize.new.parse(node_child)
        when 'tblCaption'
          @caption = ValuedChild.new(:string, parent: self).parse(node_child)
        when 'tblDescription'
          @description = ValuedChild.new(:string, parent: self).parse(node_child)
        end
      end
      @table_look = TableLook.new(parent: self).parse(node) if @table_look.nil?
      self
    end
  end
end
