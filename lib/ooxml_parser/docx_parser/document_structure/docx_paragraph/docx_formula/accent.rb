# frozen_string_literal: true

module OoxmlParser
  # Class for `acc` data
  class Accent < OOXMLDocumentObject
    # @return [ValuedChild] symbol object
    attr_reader :symbol_object
    attr_accessor :element

    # Parse Accent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Accent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'accPr'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'chr'
              @symbol_object = ValuedChild.new(:string, parent: self).parse(node_child_child)
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end

    # @return [String] symbol value
    def symbol
      @symbol_object.value
    end
  end
end
