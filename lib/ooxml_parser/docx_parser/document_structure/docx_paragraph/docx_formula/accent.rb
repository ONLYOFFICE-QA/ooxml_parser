# frozen_string_literal: true

module OoxmlParser
  # Class for `acc` data
  class Accent < OOXMLDocumentObject
    attr_accessor :symbol, :element

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
              @symbol = node_child_child.attribute('val').value
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
