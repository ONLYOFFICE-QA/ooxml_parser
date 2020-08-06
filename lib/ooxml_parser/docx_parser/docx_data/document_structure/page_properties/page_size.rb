# frozen_string_literal: true

module OoxmlParser
  # Class for data of PageSize
  class PageSize < OOXMLDocumentObject
    attr_accessor :height, :width, :orientation

    def initialize(height = nil, width = nil, orientation = :portrait)
      @height = height
      @width = width
      @orientation = orientation
    end

    # @return [String] convert to string
    def to_s
      "Height: #{@height} Width: #{@width} Orientation: #{@orientation}"
    end

    # @return [True, False] compare dimensions of size, ignoring orientation
    def same_dimensions?(other)
      (@height == other.height) && (@width == other.width) ||
        (@height == other.width) && (@width == other.height)
    end

    # @return [String] get human format name
    def name
      return 'US Letter' if same_dimensions?(PageSize.new(OoxmlSize.new(27.94, :centimeter), OoxmlSize.new(21.59, :centimeter)))
      return 'US Legal' if same_dimensions?(PageSize.new(OoxmlSize.new(35.56, :centimeter), OoxmlSize.new(21.59, :centimeter)))
      return 'A4' if same_dimensions?(PageSize.new(OoxmlSize.new(29.7, :centimeter), OoxmlSize.new(21.0, :centimeter)))
      return 'A5' if same_dimensions?(PageSize.new(OoxmlSize.new(20.99, :centimeter), OoxmlSize.new(14.81, :centimeter)))
      return 'B5' if same_dimensions?(PageSize.new(OoxmlSize.new(25.01, :centimeter), OoxmlSize.new(17.6, :centimeter)))
      return 'Envelope #10' if same_dimensions?(PageSize.new(OoxmlSize.new(24.13, :centimeter), OoxmlSize.new(10.48, :centimeter)))
      return 'Envelope DL' if same_dimensions?(PageSize.new(OoxmlSize.new(22.01, :centimeter), OoxmlSize.new(11.01, :centimeter)))
      return 'Tabloid' if same_dimensions?(PageSize.new(OoxmlSize.new(43.17, :centimeter), OoxmlSize.new(27.94, :centimeter)))
      return 'A3' if same_dimensions?(PageSize.new(OoxmlSize.new(42.01, :centimeter), OoxmlSize.new(29.7, :centimeter)))
      return 'Tabloid Oversize' if same_dimensions?(PageSize.new(OoxmlSize.new(45.71, :centimeter), OoxmlSize.new(30.48, :centimeter)))
      return 'ROC 16K' if same_dimensions?(PageSize.new(OoxmlSize.new(27.3, :centimeter), OoxmlSize.new(19.68, :centimeter)))
      return 'Envelope Choukei 3' if same_dimensions?(PageSize.new(OoxmlSize.new(23.49, :centimeter), OoxmlSize.new(11.99, :centimeter)))
      return 'Super B/A3' if same_dimensions?(PageSize.new(OoxmlSize.new(48.25, :centimeter), OoxmlSize.new(33.02, :centimeter)))

      "Unknown page size: Height #{@height} Width #{@width}"
    end

    # Parse PageSize
    # @param [Nokogiri::XML:Element] node with PageSize
    # @return [PageSize] value of Size
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'orient'
          @orientation = value.value.to_sym
        when 'h'
          @height = OoxmlSize.new(value.value.to_f)
        when 'w'
          @width = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
