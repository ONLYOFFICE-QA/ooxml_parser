# frozen_string_literal: true

module OoxmlParser
  # Class for `bar` data
  class Bar < OOXMLDocumentObject
    attr_accessor :position, :element

    def initialize(parent: nil)
      @position = :bottom
      super
    end

    # Parse Bar object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Bar] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'barPr'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'pos'
              @position = node_child_child.attribute('val').value.to_sym
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
