# frozen_string_literal: true

module OoxmlParser
  # Class for `groupChr` data
  class GroupChar < OOXMLDocumentObject
    attr_accessor :symbol, :position, :vertical_align, :element

    # Parse GroupChar object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [GroupChar] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'groupChrPr'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'chr'
              @symbol = node_child_child.attribute('m:val').value
            when 'pos'
              @position = node_child_child.attribute('m:val').value
            when 'vertJc'
              @vertical_align = node_child_child.attribute('m:val').value
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
