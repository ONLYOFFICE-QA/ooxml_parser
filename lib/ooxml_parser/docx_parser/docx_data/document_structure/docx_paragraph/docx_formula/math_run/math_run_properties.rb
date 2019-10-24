# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `m:rPr` object
  class MathRunProperties < OOXMLDocumentObject
    # @return [True, False] is run with break
    attr_accessor :break

    def initialize(parent: nil)
      @break = false
      @parent = parent
    end

    # Parse MathRunProperties
    # @param [Nokogiri::XML:Node] node with MathRunProperties
    # @return [MathRunProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |math_run_child|
        case math_run_child.name
        when 'brk'
          @break = true
        end
      end
      self
    end
  end
end
