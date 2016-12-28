require_relative 'slide/presentation_alternate_content'
require_relative 'slide/background'
require_relative 'slide/graphic_frame/graphic_frame'
require_relative 'slide/slide/shapes_grouping'
require_relative 'slide/slide/timing'
require_relative 'slide/slide_helper'
module OoxmlParser
  # Class for parsing `slide.xml`
  class Slide < OOXMLDocumentObject
    include SlideHelper
    attr_accessor :elements, :background, :transition, :timing, :alternate_content

    def initialize(parent: nil, xml_path: nil)
      @elements = []
      @background = nil
      @parent = parent
      @xml_path = xml_path
    end

    def with_data?
      return true unless @background.nil?
      @elements.each do |current_element|
        return true if current_element.with_data?
      end
      false
    end

    # Parse Slide object
    # @return [Slide] result of parsing
    def parse
      OOXMLDocumentObject.add_to_xmls_stack(@xml_path)
      node = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      node.xpath('//p:sld/*').each do |node_child|
        case node_child.name
        when 'cSld'
          node_child.xpath('*').each do |common_slide_node_child|
            case common_slide_node_child.name
            when 'spTree'
              common_slide_node_child.xpath('*').each do |slide_element_node|
                case slide_element_node.name
                when 'sp'
                  @elements << DocxShape.new(parent: self).parse(slide_element_node).dup
                when 'pic'
                  @elements << DocxPicture.new(parent: self).parse(slide_element_node)
                when 'graphicFrame'
                  @elements << GraphicFrame.new(parent: self).parse(slide_element_node)
                when 'grpSp'
                  @elements << ShapesGrouping.new(parent: self).parse(slide_element_node)
                end
              end
            when 'bg'
              @background = Background.new(parent: self).parse(common_slide_node_child)
            end
          end
        when 'timing'
          @timing = Timing.new(parent: self).parse(node_child)
        when 'transition'
          @transition = Transition.new(parent: self).parse(node_child)
        when 'AlternateContent'
          @alternate_content = PresentationAlternateContent.new(parent: self).parse(node_child)
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      self
    end
  end
end
