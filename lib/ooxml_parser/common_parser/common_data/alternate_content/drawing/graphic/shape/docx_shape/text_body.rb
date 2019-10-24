# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `txBody` tags
  class TextBody < OOXMLDocumentObject
    attr_accessor :properties, :paragraphs

    def initialize(parent: nil)
      @paragraphs = []
      @parent = parent
    end

    # Parse TextBody object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextBody] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'p'
          @paragraphs << Paragraph.new(parent: self).parse(node_child)
        when 'bodyPr'
          @properties = OOXMLShapeBodyProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
