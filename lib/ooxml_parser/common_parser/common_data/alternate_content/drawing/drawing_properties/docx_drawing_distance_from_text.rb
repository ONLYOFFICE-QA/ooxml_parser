module OoxmlParser
  # Docx Drawing Distance From Text
  class DocxDrawingDistanceFromText < OOXMLDocumentObject
    attr_accessor :top, :bottom, :left, :right

    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'distT'
          @top = OoxmlSize.new(value.value.to_f, :emu)
        when 'distB'
          @bottom = OoxmlSize.new(value.value.to_f, :emu)
        when 'distL'
          @left = OoxmlSize.new(value.value.to_f, :emu)
        when 'distR'
          @right = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      self
    end
  end
end
