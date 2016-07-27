require_relative 'numbering_definition/abstract_numbering_id'
module OoxmlParser
  # This element specifies a unique instance of numbering information that can be referenced by zero or more
  # paragraphs within the parent WordprocessingML document.
  class NumberingDefinition
    # @return [Integer] num id
    attr_accessor :id
    # @return [AbstractNumberingId] abstract numbering id
    attr_accessor :abstract_numbering_id

    # Parse NumberingDefinition data
    # @param [Nokogiri::XML:Element] node with NumberingDefinition data
    # @return [NumberingDefinition] value of Abstract Numbering data
    def self.parse(node)
      num = NumberingDefinition.new

      node.attributes.each do |key, value|
        case key
        when 'numId'
          num.id = value.value.to_f
        end
      end

      node.xpath('*').each do |numbering_child_node|
        case numbering_child_node.name
        when 'abstractNumId'
          num.abstract_numbering_id = AbstractNumberingId.parse(numbering_child_node)
        end
      end
      num
    end
  end
end
