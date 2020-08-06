# frozen_string_literal: true

require_relative 'slide/presentation_alternate_content'
require_relative 'slide/background'
require_relative 'slide/common_slide_data'
require_relative 'slide/connection_shape'
require_relative 'slide/presentation_notes'
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
    # @return [Relationships] relationships of slide
    attr_reader :relationships
    # @return [String] name of slide
    attr_reader :name
    # @return [Notes] note of slide
    attr_reader :note

    def initialize(parent: nil, xml_path: nil)
      @xml_path = xml_path
      super(parent: parent)
    end

    # @return [True, False] is slide with data
    def with_data?
      return true unless background.nil?

      elements.each do |current_element|
        return true if current_element.with_data?
      end
      false
    end

    # @return <Array> List of elements on slide
    def elements
      @common_slide_data.shape_tree.elements
    end

    # @return [Background] background of slide
    def background
      @common_slide_data.background
    end

    # Parse Slide object
    # @return [Slide] result of parsing
    def parse
      OOXMLDocumentObject.add_to_xmls_stack(@xml_path)
      @name = File.basename(@xml_path, '.*')
      node = parse_xml(OOXMLDocumentObject.current_xml)
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
      @relationships = Relationships.new(parent: self).parse_file("#{OOXMLDocumentObject.path_to_folder}#{File.dirname(@xml_path)}/_rels/#{@name}.xml.rels")
      parse_note
      self
    end

    private

    # Parse slide notes if present
    def parse_note
      notes_target = @relationships.target_by_type('notes')
      return nil if notes_target.empty?

      @note = PresentationNotes.new(parent: self).parse("#{OOXMLDocumentObject.path_to_folder}#{File.dirname(@xml_path)}/#{notes_target.first}")
    end
  end
end
