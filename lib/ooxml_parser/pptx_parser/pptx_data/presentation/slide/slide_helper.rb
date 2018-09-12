module OoxmlParser
  # Methods to help working with slide data
  module SlideHelper
    def nonempty_elements
      elements.reject { |cur_shape| cur_shape.text_body.paragraphs.first.characters.empty? }
    end

    def graphic_frames
      elements.select { |cur_element| cur_element.is_a?(GraphicFrame) }
    end

    # Get transform property of object, by object type
    # @param object [Symbol] type of object: :image, :chart, :table, :shape
    # @return [OOXMLDocumentObject] needed object
    def transform_of_object(object)
      case object
      when :image
        elements.find { |e| e.is_a? Picture }.properties.transform
      when :chart
        elements.find { |e| e.is_a? GraphicFrame }.transform
      when :table
        elements.find { |e| e.is_a? GraphicFrame }.transform
      when :shape
        elements.find { |e| !e.shape_properties.preset.nil? }.shape_properties.transform
      else
        raise "Dont know this type object - #{object}"
      end
    end

    def content_horizontal_align(object, slide_size)
      transform = transform_of_object(object)
      return :left if transform.offset.x.zero?
      return :center if OoxmlSize.new((slide_size.width.value / 2) - (transform.extents.x.value / 2)) == OoxmlSize.new(transform.offset.x.value)
      return :right if OoxmlSize.new((slide_size.width.value - transform.extents.x.value)) == OoxmlSize.new(transform.offset.x.value)

      :unknown
    end

    def content_vertical_align(object, slide_size)
      transform = transform_of_object(object)
      return :top if transform.offset.y.zero?
      return :middle if OoxmlSize.new((slide_size.height.value / 2) - (transform.extents.y.value / 2)) == OoxmlSize.new(transform.offset.y.value)
      return :bottom if OoxmlSize.new(slide_size.height.value - transform.extents.y.value) == OoxmlSize.new(transform.offset.y.value)

      :unknown
    end

    def content_distribute(object, slide_size)
      return %i[horizontally vertically] if content_horizontal_align(object, slide_size) == :center && content_vertical_align(object, slide_size) == :middle
      return [:horizontally] if content_horizontal_align(object, slide_size) == :center
      return [:vertically] if content_vertical_align(object, slide_size) == :middle
    end
  end
end
