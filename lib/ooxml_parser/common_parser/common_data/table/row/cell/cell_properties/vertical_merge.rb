module OoxmlParser
  # Class for parsing `w:vMerge` object
  class VerticalMerge < OOXMLDocumentObject
    # @return [String] value of vertical merge
    attr_accessor :value

    # Parse vMerge
    # @param [Nokogiri::XML:Node] node with vMerge
    # @return [vMerge] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_sym
        end
      end
      self
    end
  end
end
