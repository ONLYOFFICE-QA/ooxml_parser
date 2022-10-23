# frozen_string_literal: true

module OoxmlParser
  # Class for `bar` data
  class Bar < OOXMLDocumentObject
    # @return [ValuedChild] position object
    attr_reader :position_object
    attr_accessor :element

    def initialize(parent: nil)
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
              @position_object = ValuedChild.new(:symbol, parent: self).parse(node_child_child)
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end

    # @return [Symbol] position of bar
    def position
      @position_object.value
    end
  end
end
