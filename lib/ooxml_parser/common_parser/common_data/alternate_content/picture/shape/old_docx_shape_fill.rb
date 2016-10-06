# Fallback DOCX shape fill properties
module OoxmlParser
  class OldDocxShapeFill < OOXMLDocumentObject
    attr_accessor :stretching_type, :opacity, :title
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def self.parse(fill_node)
      fill = OldDocxShapeFill.new
      fill_node.attributes.each do |key, value|
        case key
        when 'id'
          fill.file_reference = FileReference.new(parent: fill).parse(fill_node)
        when 'type'
          fill.stretching_type = case value.value
                                 when 'frame'
                                   :stretch
                                 else
                                   value.value.to_sym
                                 end
        when 'title'
          fill.title = value.value.to_s
        end
      end
      fill
    end
  end
end
