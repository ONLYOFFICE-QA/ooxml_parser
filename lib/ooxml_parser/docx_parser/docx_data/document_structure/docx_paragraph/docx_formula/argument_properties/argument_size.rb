module OoxmlParser
  # Class for parsing `m:argSz` object
  class ArgumentSize
    # @return [String] value of size
    attr_accessor :value

    # Parse ArgumentSize
    # @param [Nokogiri::XML:Node] node with ArgumentSize
    # @return [ArgumentSize] result of parsing
    def self.parse(node)
      argument_size = ArgumentSize.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          argument_size.value = value.value.to_f
        end
      end
      argument_size
    end
  end
end
