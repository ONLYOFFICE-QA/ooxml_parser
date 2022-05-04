# frozen_string_literal: true

module OoxmlParser
  # Class for 'f' data
  class Fraction < OOXMLDocumentObject
    attr_accessor :numerator, :denominator

    # Parse Fraction object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Fraction] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'num'
          @numerator = DocxFormula.new(parent: self).parse(node_child)
        when 'den'
          @denominator = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
