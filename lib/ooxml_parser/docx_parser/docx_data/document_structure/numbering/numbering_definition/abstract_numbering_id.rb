module OoxmlParser
  # Class for storing AbstractNumberingId
  class AbstractNumberingId
    # @return [String] value of start
    attr_accessor :value

    # Parse AbstractNumberingId
    # @param [Nokogiri::XML:Node] node with AbstractNumberingId
    # @return [AbstractNumberingId] result of parsing
    def self.parse(node)
      abstract_id = AbstractNumberingId.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          abstract_id.value = value.value.to_i
        end
      end
      abstract_id
    end
  end
end
