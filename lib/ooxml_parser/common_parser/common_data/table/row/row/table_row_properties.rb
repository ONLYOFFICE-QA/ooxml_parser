module OoxmlParser
  # Class for describing Table Row Properties
  class TableRowProperties < OOXMLDocumentObject
    # @return [Float] Table Row Height
    attr_accessor :height

    def initialize
      @height = nil
    end

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Table Row Properties data
    # @return [TableRowProperties] value of Columns data
    def self.parse(node)
      properties = TableRowProperties.new
      node.xpath('w:trHeight').each do |height|
        properties.height = height.attribute('val').value.to_f
      end
      properties
    end
  end
end
