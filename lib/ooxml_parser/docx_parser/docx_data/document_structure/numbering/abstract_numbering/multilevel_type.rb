module OoxmlParser
  # This element specifies the type of numbering defined by a given abstract numbering type
  class MultilevelType
    # @return [String] value of Multilevel Type
    attr_accessor :value

    # Parse MultilevelType
    # @param [Nokogiri::XML:Node] node with MultilevelType
    # @return [MultilevelType] result of parsing
    def self.parse(node)
      multilevel_type = MultilevelType.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          multilevel_type.value = Alignment.parse(value)
        end
      end
      multilevel_type
    end
  end
end
