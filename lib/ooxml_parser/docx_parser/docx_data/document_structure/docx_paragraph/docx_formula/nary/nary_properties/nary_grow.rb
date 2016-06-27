module OoxmlParser
  # Class for parsing `m:grow` object
  class NaryGrow
    # @return [String] value of grow
    attr_accessor :value

    # Parse NaryGrow
    # @param [Nokogiri::XML:Node] node with NaryGrow
    # @return [NaryGrow] result of parsing
    def self.parse(node)
      nary_grow = NaryGrow.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          nary_grow.value = value.value
        end
      end
      nary_grow
    end
  end
end
