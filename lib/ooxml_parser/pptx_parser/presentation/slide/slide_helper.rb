# frozen_string_literal: true

module OoxmlParser
  # Methods to help working with slide data
  module SlideHelper
    # @return [Array] list of not empty element on slide
    def nonempty_elements
      elements.reject { |cur_shape| cur_shape.text_body.paragraphs.first.characters.empty? }
    end

    # @return [Array<GraphicFrame>] list GraphicFrame elements on slide
    def graphic_frames
      elements.select { |cur_element| cur_element.is_a?(GraphicFrame) }
    end

    # Get transform property of object, by object type
    # @param object [Symbol] type of object: :image, :chart, :table, :shape
    # @return [OOXMLDocumentObject] needed object
    def transform_of_object(object)
      case object
      when :image
        elements.find { |e| e.is_a? DocxPicture }.properties.transform
      when :chart, :table
        elements.find { |e| e.is_a? GraphicFrame }.transform
      when :shape
        elements.find { |e| !e.shape_properties.preset.nil? }.shape_properties.transform
      else
        raise "Dont know this type object - #{object}"
      end
    end

    # Get horizontal align of object on slide
    # @param object [Symbol] object to get
    # @param slide_size [SlideSize] size of slide
    # @return [Symbol] type of align
    def content_horizontal_align(object, slide_size)
      transform = transform_of_object(object)
      return :left if transform.offset.x.zero?
      return :center if OoxmlSize.new((slide_size.width.value / 2) - (transform.extents.x.value / 2)) == OoxmlSize.new(transform.offset.x.value)
      return :right if OoxmlSize.new((slide_size.width.value - transform.extents.x.value)) == OoxmlSize.new(transform.offset.x.value)

      :unknown
    end

    # Get vertical align of object on slide
    # @param object [Symbol] object to get
    # @param slide_size [SlideSize] size of slide
    # @return [Symbol] type of align
    def content_vertical_align(object, slide_size)
      transform = transform_of_object(object)
      return :top if transform.offset.y.zero?
      return :middle if OoxmlSize.new((slide_size.height.value / 2) - (transform.extents.y.value / 2)) == OoxmlSize.new(transform.offset.y.value)
      return :bottom if OoxmlSize.new(slide_size.height.value - transform.extents.y.value) == OoxmlSize.new(transform.offset.y.value)

      :unknown
    end

    # Get content distribution of object
    # @param object [Symbol] object to get
    # @param slide_size [SlideSize] size of slide
    # @return [Array<Symbol>] type of align
    def content_distribute(object, slide_size)
      return %i[horizontally vertically] if content_horizontal_align(object, slide_size) == :center && content_vertical_align(object, slide_size) == :middle
      return [:horizontally] if content_horizontal_align(object, slide_size) == :center

      [:vertically] if content_vertical_align(object, slide_size) == :middle
    end
  end
end
