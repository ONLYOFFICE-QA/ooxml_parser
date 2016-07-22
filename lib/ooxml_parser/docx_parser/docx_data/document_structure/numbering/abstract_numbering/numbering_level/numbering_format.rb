module OoxmlParser
  # Class for storing Numbering Format data
  class NumberingFormat
    # @return [String] value of start
    attr_accessor :value

    # Parse NumberingFormat
    # @param [Nokogiri::XML:Node] node with NumberingFormat
    # @return [NumberingFormat] result of parsing
    def self.parse(node)
      num_format = NumberingFormat.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          num_format.value = value.value.to_s
        end
      end
      num_format
    end
  end
end
