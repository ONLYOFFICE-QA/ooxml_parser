module OoxmlParser
  # This element specifies a unique instance of numbering information that can be referenced by zero or more
  # paragraphs within the parent WordprocessingML document.
  class NumberingDefinition < OOXMLDocumentObject
    # @return [Integer] num id
    attr_accessor :id
    # @return [ValuedChild] abstract numbering id
    attr_accessor :abstract_numbering_id

    # Parse NumberingDefinition data
    # @param [Nokogiri::XML:Element] node with NumberingDefinition data
    # @return [NumberingDefinition] value of Abstract Numbering data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'numId'
          @id = value.value.to_f
        end
      end

      node.xpath('*').each do |numbering_child_node|
        case numbering_child_node.name
        when 'abstractNumId'
          @abstract_numbering_id = ValuedChild.new(:integer, parent: self).parse(numbering_child_node)
        end
      end
      self
    end
  end
end
