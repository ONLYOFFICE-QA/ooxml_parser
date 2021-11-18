# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `hlinkClick`, `hyperlink` tags
  class Hyperlink < OOXMLDocumentObject
    # @return [OOXMLDocumentObject] url of hyperlink
    attr_accessor :url
    # @return [String] tooltip value
    attr_reader :tooltip
    # @return [Coordinates] coordinates of link
    attr_reader :coordinates
    # @return [True, False] should click be highlighted
    attr_reader :highlight_click
    # @return [Symbol] type of action
    attr_reader :action
    # @return [String] value of action link
    attr_reader :action_link
    # @return [String] id of link
    attr_reader :id
    # @return [Array<ParagraphRun>] run of paragraph
    attr_reader :runs

    def initialize(link = nil,
                   tooltip = nil,
                   coordinates = nil,
                   parent: nil)
      @url = link
      @tooltip = tooltip
      @coordinates = coordinates
      @runs = []
      super(parent: parent)
    end

    alias link url
    alias link_to url
    alias link= url=
    alias link_to= url=

    # Parse Hyperlink object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Hyperlink] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'location'
          @url = Coordinates.parse_coordinates_from_string(value.value)
        when 'id'
          @id = value.value
          @url = OOXMLDocumentObject.get_link_from_rels(@id) unless @id.empty?
        when 'tooltip'
          @tooltip = value.value
        when 'ref'
          @coordinates = Coordinates.parse_coordinates_from_string(value.value)
        when 'action'
          @action_link = value.value
        when 'highlightClick'
          @highlight_click = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'r'
          @runs << ParagraphRun.new(parent: self).parse(node_child)
        end
      end

      case @action_link
      when 'ppaction://hlinkshowjump?jump=previousslide'
        @action = :previous_slide
      when 'ppaction://hlinkshowjump?jump=nextslide'
        @action = :next_slide
      when 'ppaction://hlinkshowjump?jump=firstslide'
        @action = :first_slide
      when 'ppaction://hlinkshowjump?jump=lastslide'
        @action = :last_slide
      when 'ppaction://hlinksldjump'
        @action = :slide
        parse_url_for_slide_link
      else
        if meaningful_id?
          @action = :external_link
          @url = OOXMLDocumentObject.get_link_from_rels(@id)
        end
      end
      self
    end

    private

    # Check if id parameter has any information in it
    # @return [Boolean] Can id be used
    def meaningful_id?
      @id && !@id.empty?
    end

    # Parse url for slide link
    def parse_url_for_slide_link
      return unless meaningful_id?

      @url = OOXMLDocumentObject.get_link_from_rels(@id).scan(/\d+/).join.to_i
    end
  end
end
