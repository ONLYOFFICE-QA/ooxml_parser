module OoxmlParser
  # Class for parsing `w:bookmarkStart`, `w:bookmarkEnd` tags
  class Bookmark < OOXMLDocumentObject
    attr_accessor :id, :name

    # Parse Bookmark object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Bookmark] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value
        when 'name'
          @name = value.value
        end
      end
      self
    end
  end
end
