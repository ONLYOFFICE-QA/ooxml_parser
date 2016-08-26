module OoxmlParser
  class PageSize < OOXMLDocumentObject
    attr_accessor :height, :width, :orientation

    def initialize(height = nil, width = nil, orientation = :portrait)
      @height = height
      @width = width
      @orientation = orientation
    end

    def self.get_name_by_size(height, width, is_portrait = true)
      unless is_portrait
        swap_value = height
        height = width
        width = swap_value
      end
      return 'US Letter' if height == '15840' && width == '12240'
      return 'US Legal' if height == '20160' && width == '12240'
      return 'A4' if height == '16838' && width == '11906'
      return 'A5' if height == '11900' && width == '8396'
      return 'B5' if height == '14179' && width == '9978'
      return 'Envelope #10' if height == '13680' && width == '5941'
      return 'Envelope DL' if height == '12478' && width == '6242'
      return 'Tabloid' if height == '24474' && width == '15840'
      return 'A3' if height == '23817' && width == '16838'
      return 'Tabloid Oversize' if height == '25914' && width == '17280'
      return 'ROC 16K' if height == '15477' && width == '11157'
      return 'Envelope Choukei 3' if height == '13317' && width == '6797'
      return 'Super B/A3' if height == '27354' && width == '18720'
      'Unknown page size: Height ' + height.to_s + ' Width ' + width.to_s
    end

    def to_s
      'Height: ' + @height.to_s + ' Width: ' + @width.to_s + ' Orientation: ' + @orientation.to_s
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
