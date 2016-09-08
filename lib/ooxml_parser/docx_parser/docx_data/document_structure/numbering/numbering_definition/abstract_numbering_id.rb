module OoxmlParser
  # Class for storing AbstractNumberingId
  class AbstractNumberingId < OOXMLDocumentObject
    # @return [String] value of start
    attr_accessor :value

    # Parse AbstractNumberingId
    # @param [Nokogiri::XML:Node] node with AbstractNumberingId
    # @return [AbstractNumberingId] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_i
        end
      end
      self
    end
  end
end
