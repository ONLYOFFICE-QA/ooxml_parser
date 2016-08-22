module OoxmlParser
  class SlideSize
    attr_accessor :width, :height, :type

    def self.parse(slide_size_node)
      slide_size = SlideSize.new
      slide_size_node.attributes.each do |key, value|
        case key
        when 'cx'
          slide_size.width = OoxmlSize.new(value.value.to_f, :emu)
        when 'cy'
          slide_size.height = OoxmlSize.new(value.value.to_f, :emu)
        when 'type'
          slide_size.type = value.value.to_sym
        end
      end
      slide_size
    end
  end
end
