require_relative 'slide/presentation_alternate_content'
require_relative 'slide/background'
require_relative 'slide/common_slide_data'
require_relative 'slide/connection_shape.rb'
require_relative 'slide/graphic_frame/graphic_frame'
require_relative 'slide/slide/shapes_grouping'
require_relative 'slide/slide/timing'
require_relative 'slide/slide_helper'
module OoxmlParser
  # Class for parsing `slide.xml`
  class Slide < OOXMLDocumentObject
    include SlideHelper
    attr_accessor :transition, :timing, :alternate_content
    # @return [CommonSlideData] common slide data
    attr_reader :common_slide_data

    def initialize(parent: nil, xml_path: nil)
      @parent = parent
      @xml_path = xml_path
    end

    def with_data?
      return true unless background.nil?
      elements.each do |current_element|
        return true if current_element.with_data?
      end
      false
    end

    def elements
      @common_slide_data.shape_tree.elements
    end

    def background
      @common_slide_data.background
    end

    # Parse Slide object
    # @return [Slide] result of parsing
    def parse
      OOXMLDocumentObject.add_to_xmls_stack(@xml_path)
      node = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      node.xpath('//p:sld/*').each do |node_child|
        case node_child.name
        when 'cSld'
          @common_slide_data = CommonSlideData.new(parent: self).parse(node_child)
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
