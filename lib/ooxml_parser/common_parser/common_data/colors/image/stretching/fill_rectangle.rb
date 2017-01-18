module OoxmlParser
  # Class for parsing `fillRect` tag
  class FillRectangle < OOXMLDocumentObject
    attr_accessor :left, :top, :right, :bottom

    # Parse FillRectangle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FillRectangle] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'b'
          @bottom = value.value.to_i
        when 't'
          @top = value.value.to_i
        when 'l'
          @left = value.value.to_i
        when 'r'
          @right = value.value.to_i
        end
      end
      self
    end
  end
end
