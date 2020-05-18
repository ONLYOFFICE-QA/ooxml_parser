# frozen_string_literal: true

require_relative 'math_text/math_paragraph'
module OoxmlParser
  # Class for storing math text
  class MathText < OOXMLDocumentObject
    attr_accessor :math_paragraph

    # Parse MathText object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [MathText] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'oMathPara'
          @math_paragraph = MathParagraph.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
