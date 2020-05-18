# frozen_string_literal: true

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
    # @return [TableProperties] properties of table
    attr_accessor :table_properties
    # @return [TableRowProperties] properties of table row
    attr_reader :table_row_properties
    # @return [ParagraphProperties] properties of paragraph
    attr_accessor :paragraph_properties

    alias cell_properties table_cell_properties

    def initialize(type: nil, parent: nil)
      @type = type
      @run_properties = nil
      @table_cell_properties = CellProperties.new
      @parent = parent
    end

    # Parse table style property
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [TableStyleProperties]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          @run_properties = RunProperties.new(parent: self).parse(node_child)
        when 'tcPr'
          @table_cell_properties = CellProperties.new(parent: self).parse(node_child)
        when 'tblPr'
          @table_properties = TableProperties.new(parent: self).parse(node_child)
        when 'trPr'
          @table_row_properties = TableRowProperties.new(parent: self).parse(node_child)
        when 'pPr'
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
