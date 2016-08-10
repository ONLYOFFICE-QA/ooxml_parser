# Table Properties Data
require_relative 'properties/table_look'
require_relative 'properties/table_style'
require_relative 'properties/table_position'
require_relative 'properties/table_style_properties'
require_relative 'table_properties/table_borders'
module OoxmlParser
  class TableProperties < OOXMLDocumentObject
    attr_accessor :jc, :table_width, :shd, :table_borders, :table_properties, :table_positon, :table_cell_margin, :table_indent, :stretching, :table_style, :row_banding_size,
                  :column_banding_size, :table_look, :grid_column, :right_to_left, :style

    alias table_properties table_positon

    def initialize
      @table_indent = nil
      @table_cell_margin = nil
      @jc = :left
      @shd = nil
      @stretching = true
      @table_borders = TableBorders.new
      @table_properties = nil
      @table_width = nil
      @grid_column = nil
      @right_to_left = nil
      @style = nil
    end

    def copy
      table = TableProperties.new
      table.jc = @jc
      table.table_width = @table_width
      table.shd = @shd
      table.stretching = @stretching
      table.table_borders = @table_borders
      table.table_properties = @table_properties
      table.table_cell_margin = @table_cell_margin
      table.table_indent = @table_indent
      table.grid_column = @grid_column
      table.right_to_left = @right_to_left
      table.style = style
      table
    end

    def self.parse(table_properties_node, parent: nil)
      table_properties = TableProperties.new
      table_properties.parent = parent
      table_properties_node.xpath('*').each do |table_props_node_child|
        case table_props_node_child.name
        when 'tableStyleId'
          table_properties.style = TableStyle.parse(style_id: table_props_node_child.text)
        when 'tblBorders'
          table_properties.table_borders = TableBorders.parse(table_props_node_child)
        when 'tblStyle'
          table_properties.table_style = table_properties.root_object.document_style_by_id(table_props_node_child.attribute('val').value)
        when 'tblW'
          table_properties.table_width = table_props_node_child.attribute('w').text.to_f / 567.0
        when 'jc'
          table_properties.jc = table_props_node_child.attribute('val').text.to_sym
        when 'shd'
          unless table_props_node_child.attribute('fill').nil?
            background_color = Color.from_int16(table_props_node_child.attribute('fill').value)
            table_properties.shd = background_color
          end
        when 'tblLook'
          table_properties.table_look = TableLook.parse(table_props_node_child)
        when 'tblInd'
          table_properties.table_indent = table_props_node_child.attribute('w').text.to_f / 567.0
        when 'tblpPr'
          table_properties.table_positon = TablePosition.parse(table_props_node_child)
        when 'tblCellMar'
          table_properties.table_cell_margin = TableMargins.parse(table_props_node_child)
        end
      end
      table_properties.table_look = TableLook.parse(table_properties_node) if table_properties.table_look.nil?
      table_properties.right_to_left = OOXMLDocumentObject.option_enabled?(table_properties_node, 'rtl')
      table_properties
    end
  end
end
