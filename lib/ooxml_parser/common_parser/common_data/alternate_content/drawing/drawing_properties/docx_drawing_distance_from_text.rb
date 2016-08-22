# Docx Drawing Distance From Text
module OoxmlParser
  class DocxDrawingDistanceFromText
    attr_accessor :top, :bottom, :left, :right

    def self.parse(distance_node)
      distance_from_text = DocxDrawingDistanceFromText.new
      distance_node.attributes.each do |key, value|
        case key
        when 'distT'
          distance_from_text.top = OoxmlSize.new(value.value.to_f, :emu)
        when 'distB'
          distance_from_text.bottom = OoxmlSize.new(value.value.to_f, :emu)
        when 'distL'
          distance_from_text.left = OoxmlSize.new(value.value.to_f, :emu)
        when 'distR'
          distance_from_text.right = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      distance_from_text
    end
  end
end
