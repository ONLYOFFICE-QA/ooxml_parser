# Docx Drawing Distance From Text
module OoxmlParser
  class DocxDrawingDistanceFromText
    attr_accessor :top, :bottom, :left, :right

    def self.parse(distance_node)
      distance_from_text = DocxDrawingDistanceFromText.new
      distance_node.attributes.each do |key, value|
        case key
        when 'distT'
          distance_from_text.top = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        when 'distB'
          distance_from_text.bottom = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        when 'distL'
          distance_from_text.left = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        when 'distR'
          distance_from_text.right = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        end
      end
      distance_from_text
    end
  end
end
