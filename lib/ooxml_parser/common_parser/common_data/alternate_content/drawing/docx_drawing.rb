# frozen_string_literal: true

require_relative 'docx_drawing/doc_properties'
require_relative 'drawing_properties/docx_drawing_properties'
require_relative 'drawing_properties/size_relative_horizontal'
require_relative 'drawing_properties/size_relative_vertical'
require_relative 'graphic/docx_graphic'
module OoxmlParser
  # Class for parsing `graphic` tags
  class DocxDrawing < OOXMLDocumentObject
    # @return [String] id of drawing
    attr_reader :id
    attr_accessor :type, :properties, :graphic
    # @return [DocProperties] doc properties
    attr_accessor :doc_properties

    alias picture graphic

    def initialize(properties = DocxDrawingProperties.new, parent: nil)
      @properties = properties
      super(parent: parent)
    end

    # Parse DocxDrawing
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [DocxDrawing] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'anchor'
          @type = :flow
        when 'inline'
          @type = :inline
        end
        node_child.xpath('*').each do |content_node_child|
          case content_node_child.name
          when 'graphic'
            @graphic = DocxGraphic.new(parent: self).parse(content_node_child)
          when 'docPr'
            @doc_properties = DocProperties.new(parent: self).parse(content_node_child)
          end
        end
        @properties.parse(node_child)
      end
      self
    end
  end
end
