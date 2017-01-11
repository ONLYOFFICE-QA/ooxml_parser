# Docx Wrap Drawing
module OoxmlParser
  class DocxWrapDrawing
    attr_accessor :wrap_text, :distance_from_text

    def initialize(wrap_text = :none)
      @wrap_text = wrap_text
    end

    def self.parse(drawing_node)
      wrap = DocxWrapDrawing.new
      unless drawing_node.attribute('behindDoc').nil?
        wrap.wrap_text = :behind if drawing_node.attribute('behindDoc').value == '1'
        wrap.wrap_text = :infront if drawing_node.attribute('behindDoc').value == '0'
      end
      drawing_node.xpath('*').each do |wrap_node|
        case wrap_node.name
        when 'wrapSquare'
          wrap.wrap_text = :square
          wrap.distance_from_text = DocxDrawingDistanceFromText.new(parent: wrap).parse(wrap_node)
          break
        when 'wrapTight'
          wrap.wrap_text = :tight
          wrap.distance_from_text = DocxDrawingDistanceFromText.new(parent: wrap).parse(wrap_node)
          break
        when 'wrapThrough'
          wrap.wrap_text = :through
          wrap.distance_from_text = DocxDrawingDistanceFromText.new(parent: wrap).parse(wrap_node)
          break
        when 'wrapTopAndBottom'
          wrap.wrap_text = :topbottom
          wrap.distance_from_text = DocxDrawingDistanceFromText.new(parent: wrap).parse(wrap_node)
          break
        end
      end
      wrap
    end
  end
end
