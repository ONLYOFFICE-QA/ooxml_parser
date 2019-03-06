module OoxmlParser
  # Class for parsing `formulas` <f>
  class Formula < OOXMLDocumentObject
    # @return [Coordinates] reference coordinates
    attr_reader :reference
    # @return [StringIndex] string index
    attr_reader :string_index
    # @return [String] type
    attr_reader :type
    # @return [String] value
    attr_reader :value

    # Parse Formula object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Formula] parsed object
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ref'
          @reference = Coordinates.parser_coordinates_range(value.value.to_s)
        when 'si'
          @string_index = value.value.to_i
        when 't'
          @type = value.value.to_s
        end
      end

      @value = node.text unless node.text.empty?
      self
    end

    # @return [True, False] check if formula empty
    def empty?
      !(reference || string_index || type || value)
    end
  end
end
