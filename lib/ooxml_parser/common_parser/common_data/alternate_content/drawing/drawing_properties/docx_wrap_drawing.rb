module OoxmlParser
  # Docx Wrap Drawing
  class DocxWrapDrawing < OOXMLDocumentObject
    attr_accessor :wrap_text, :distance_from_text

    def initialize(parent: nil)
      @wrap_text = :none
      @parent = parent
    end

    # Parse DocxWrapDrawing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxWrapDrawing] result of parsing
    def parse(node)
      unless node.attribute('behindDoc').nil?
        @wrap_text = :behind if node.attribute('behindDoc').value == '1'
        @wrap_text = :infront if node.attribute('behindDoc').value == '0'
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'wrapSquare'
          @wrap_text = :square
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapTight'
          @wrap_text = :tight
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapThrough'
          @wrap_text = :through
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapTopAndBottom'
          @wrap_text = :topbottom
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        end
      end
      self
    end
  end
end
