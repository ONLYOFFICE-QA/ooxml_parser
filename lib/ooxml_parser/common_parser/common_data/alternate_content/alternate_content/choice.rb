# frozen_string_literal: true

require_relative 'choice/math_text'
module OoxmlParser
  # Class for storing choice (for office 2014 and newer)
  # graphic data
  class Choice < OOXMLDocumentObject
    attr_accessor :math_text

    # Parse Choice object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Choice] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'm'
          @math_text = MathText.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
