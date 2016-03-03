require_relative 'slide/presentation_alternate_content'
require_relative 'slide/background'
require_relative 'slide/graphic_frame/graphic_frame'
require_relative 'slide/slide/shapes_grouping'
require_relative 'slide/slide/timing'
module OoxmlParser
  class Slide < OOXMLDocumentObject
    attr_accessor :elements, :background, :transition, :timing, :alternate_content

    def initialize(elements = [], background = nil)
      @elements = elements
      @background = background
    end

    def nonempty_elements
      @elements.select { |cur_shape| !cur_shape.text_body.paragraphs.first.characters.empty? }
    end

    def graphic_frames
      @elements.select { |cur_element| cur_element.is_a?(GraphicFrame) }
    end

    def content_horizontal_align(object, slide_size)
      transform = nil
      case object
      when :image
        transform = elements.find { |e| e.is_a? Picture }.properties.transform
      when :chart
        transform = elements.find { |e| e.is_a? GraphicFrame }.transform
      when :table
        transform = elements.find { |e| e.is_a? GraphicFrame }.transform
      when :shape
        transform = elements.find { |e| !e.shape_properties.preset.nil? }.shape_properties.transform
      else
        raise "Dont know this type object - #{object}"
      end
      return :left if transform.offset.x == 0
      return :center if ((slide_size.width / 2) - (transform.extents.x / 2)).round(1) == transform.offset.x.round(1)
      return :right if (slide_size.width - transform.extents.x).round(1) == transform.offset.x.round(1)
      :unknown
    end

    def content_vertical_align(object, slide_size)
      transform = nil
      case object
      when :image
        transform = elements.find { |e| e.is_a? Picture }.properties.transform
      when :chart
        transform = elements.find { |e| e.is_a? GraphicFrame }.transform
      when :table
        transform = elements.find { |e| e.is_a? GraphicFrame }.transform
      when :shape
        transform = elements.find { |e| !e.shape_properties.preset.nil? }.shape_properties.transform
      else
        raise "Dont know this type object - #{object}"
      end
      return :top if transform.offset.y == 0
      return :middle if ((slide_size.height / 2) - (transform.extents.y / 2)).round(1) == transform.offset.y.round(1)
      return :bottom if (slide_size.height - transform.extents.y).round(1) == transform.offset.y.round(1)
      :unknown
    end

    def content_distribute(object, slide_size)
      return [:horizontally, :vertically] if content_horizontal_align(object, slide_size) == :center && content_vertical_align(object, slide_size) == :middle
      return [:horizontally] if content_horizontal_align(object, slide_size) == :center
      return [:vertically] if content_vertical_align(object, slide_size) == :middle
    end

    def self.parse(path_to_slide_xml)
      slide = Slide.new
      OOXMLDocumentObject.add_to_xmls_stack(path_to_slide_xml)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      doc.xpath('//p:sld/*').each do |slide_node_child|
        case slide_node_child.name
        when 'cSld'
          slide_node_child.xpath('*').each do |common_slide_node_child|
            case common_slide_node_child.name
            when 'spTree'
              common_slide_node_child.xpath('*').each do |slide_element_node|
                case slide_element_node.name
                when 'nvGrpSpPr'
                when 'grpSpPr'
                when 'sp'
                  slide.elements << PresentationShape.parse(slide_element_node).dup
                when 'pic'
                  slide.elements << DocxPicture.parse(slide_element_node)
                when 'graphicFrame'
                  slide.elements << GraphicFrame.parse(slide_element_node)
                when 'grpSp'
                  slide.elements << ShapesGrouping.parse(slide_element_node)
                end
              end
            when 'bg'
              slide.background = Background.parse(common_slide_node_child)
            end
          end
        when 'clrMapOvr'
        when 'timing'
          slide.timing = Timing.parse(slide_node_child)
        when 'transition'
          slide.transition = Transition.parse(slide_node_child)
        when 'AlternateContent'
          slide.alternate_content = PresentationAlternateContent.parse(slide_node_child)
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      slide
    end
  end
end
