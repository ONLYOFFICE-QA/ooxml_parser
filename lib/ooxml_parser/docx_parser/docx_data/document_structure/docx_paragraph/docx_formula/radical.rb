# frozen_string_literal: true

module OoxmlParser
  # Class for `rad` data
  class Radical < OOXMLDocumentObject
    attr_accessor :degree, :value

    def initialize(parent: nil)
      @degree = 2
      super
    end

    # Parse Radical object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Radical] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'deg'
          @degree = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      @value = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
