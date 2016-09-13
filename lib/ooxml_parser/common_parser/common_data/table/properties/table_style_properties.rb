require_relative 'table_style_properties/table_style_properties_helper'
module OoxmlParser
  # Class for parsing `w:tblStylePr`
  class TableStyleProperties < OOXMLDocumentObject
    # @return [Symbol] type of Table Style Properties
    attr_accessor :type
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [CellProperties] properties of table cell
    attr_accessor :table_cell_properties
    alias cell_properties table_cell_properties

    def initialize(type: nil)
      @type = type
      @run_properties = nil
      @table_cell_properties = CellProperties.new
    end

    # Parse table style property
    # @param node [Nokogiri::XML::Element] node to parse
    # @param [OoxmlParser::OOXMLDocumentObject] parent parent object
    # @return [TableStyleProperties]
    def self.parse(node, parent: nil)
      table_style_pr = TableStyleProperties.new
      table_style_pr.parent = parent

      node.attributes.each do |key, value|
        case key
        when 'type'
          table_style_pr.type = value.value.to_sym
        end
      end

      node.xpath('*').each do |properties_child|
        case properties_child.name
        when 'rPr'
          table_style_pr.run_properties = RunProperties.new(parent: table_style_pr).parse(properties_child)
        when 'tcPr'
          table_style_pr.table_cell_properties = CellProperties.new(parent: table_style_pr).parse(properties_child)
        end
      end
      table_style_pr
    end
  end
end
