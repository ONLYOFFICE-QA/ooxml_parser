module OoxmlParser
  class DocumentGrid < OOXMLDocumentObject
    attr_accessor :type, :line_pitch, :char_space

    def initialize(type = nil, line_pitch = nil, char_space = nil)
      @type = type
      @line_pitch = line_pitch
      @char_space = char_space
    end

    # Parse DocumentGrid
    # @param [Nokogiri::XML:Element] node with DocumentGrid
    # @return [DocumentGrid] value of DocumentGrid
    def self.parse(node)
      document_grid = DocumentGrid.new
      node.attributes.each do |key, value|
        case key
        when 'charSpace'
          document_grid.char_space = value.value
        when 'linePitch'
          document_grid.line_pitch = value.value.to_i
        when 'type'
          document_grid.type = value.value
        end
      end
      document_grid
    end
  end
end
