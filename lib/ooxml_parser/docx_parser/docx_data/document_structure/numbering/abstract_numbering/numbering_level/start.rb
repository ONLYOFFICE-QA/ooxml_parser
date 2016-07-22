module OoxmlParser
  # Class for storing Start data
  class Start
    # @return [Integer] value of start
    attr_accessor :value

    # Parse Start
    # @param [Nokogiri::XML:Node] node with Start
    # @return [Start] result of parsing
    def self.parse(node)
      start = Start.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          start.value = value.value.to_f
        end
      end
      start
    end
  end
end
