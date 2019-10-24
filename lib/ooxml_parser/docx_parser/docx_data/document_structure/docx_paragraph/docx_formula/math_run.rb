# frozen_string_literal: true

require_relative 'math_run/math_run_properties'
module OoxmlParser
  # Class for parsing `m:r` object
  class MathRun < OOXMLDocumentObject
    # @return [MathRunProperties] properties of run
    attr_accessor :properties
    # @return [String] text of formula
    attr_accessor :text

    # Parse MathRun
    # @param [Nokogiri::XML:Node] node with MathRun
    # @return [MathRun] result of parsing
    def parse(node)
      node.xpath('*').each do |math_run_child|
        case math_run_child.name
        when 'rPr'
          @properties = MathRunProperties.new(parent: self).parse(math_run_child)
        when 't'
          @text = math_run_child.text
        end
      end
      self
    end
  end
end
