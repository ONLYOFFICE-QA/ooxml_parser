require_relative 'slide/presentation_alternate_content'
require_relative 'slide/background'
require_relative 'slide/graphic_frame/graphic_frame'
require_relative 'slide/slide/shapes_grouping'
require_relative 'slide/slide/timing'
require_relative 'slide/slide_helper'
module OoxmlParser
  class Slide < OOXMLDocumentObject
    include SlideHelper
    attr_accessor :elements, :background, :transition, :timing, :alternate_content

    def initialize(elements = [], background = nil)
      @elements = elements
      @background = background
    end

    def with_data?
      return true unless @background.nil?
      @elements.each do |current_element|
        return true if current_element.with_data?
      end
      false
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
              slide.background = Background.new(parent: slide).parse(common_slide_node_child)
            end
          end
        when 'timing'
          slide.timing = Timing.parse(slide_node_child)
        when 'transition'
          slide.transition = Transition.new(parent: slide).parse(slide_node_child)
        when 'AlternateContent'
          slide.alternate_content = PresentationAlternateContent.parse(slide_node_child)
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      slide
    end
  end
end
