module OoxmlParser
  # Class for parsing `w:sz` object
  class RunSpacing < OOXMLDocumentObject
    # @return [Integer] value of size
    attr_accessor :value

    # Parse RunSpacing
    # @param [Nokogiri::XML:Node] node with RunSpacing
    # @return [RunSpacing] result of parsing
    def self.parse(node)
      index = RunSpacing.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          index.value = OoxmlSize.new(value.value.to_f)
        end
      end
      index
    end
  end
end
