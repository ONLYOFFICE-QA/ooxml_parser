# frozen_string_literal: true

module OoxmlParser
  # Parsing chart `c:pivotFmt`
  class PivotFormat < OOXMLDocumentObject
    # @return [ValuedChild] index of pivot
    attr_reader :index
    # @return [DisplayLabelsProperties] data label
    attr_reader :data_label

    # Parse PivotFormats object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PivotFormats] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'idx'
          @index = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'dLbl'
          @data_label = DisplayLabelsProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
