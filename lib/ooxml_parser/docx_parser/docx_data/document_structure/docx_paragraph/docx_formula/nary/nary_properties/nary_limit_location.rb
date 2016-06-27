module OoxmlParser
  # Class for parsing `m:limLoc` object
  class NaryLimitLocation
    # @return [String] value of limit location
    attr_accessor :value

    # Parse NaryLimitLocation
    # @param [Nokogiri::XML:Node] node with NaryLimitLocation
    # @return [NaryLimitLocation] result of parsing
    def self.parse(node)
      limit_location = NaryLimitLocation.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          limit_location.value = Alignment.parse(value)
        end
      end
      limit_location
    end
  end
end
