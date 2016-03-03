module OoxmlParser
  class Size
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
      if height == '15840' && width == '12240'
        return 'US Letter'
      elsif height == '20160' && width == '12240'
        return 'US Legal'
      elsif height == '16838' && width == '11906'
        return 'A4'
      elsif height == '11900' && width == '8396'
        return 'A5'
      elsif height == '14179' && width == '9978'
        return 'B5'
      elsif height == '13680' && width == '5941'
        return 'Envelope #10'
      elsif height == '12478' && width == '6242'
        return 'Envelope DL'
      elsif height == '24474' && width == '15840'
        return 'Tabloid'
      elsif height == '23817' && width == '16838'
        return 'A3'
      elsif height == '25914' && width == '17280'
        return 'Tabloid Oversize'
      elsif height == '15477' && width == '11157'
        return 'ROC 16K'
      elsif height == '13317' && width == '6797'
        return 'Envelope Choukei 3'
      elsif height == '27354' && width == '18720'
        return 'Super B/A3'
      else
        return 'Unknown page size: Height ' + height.to_s + ' Width ' + width.to_s
      end
    end

    def to_s
      'Height: ' + @height.to_s + ' Width: ' + @width.to_s + ' Orientation: ' + @orientation.to_s
    end

    # noinspection RubyUnnecessaryReturnStatement
    def ==(other)
      is_same = false
      if (@height == other.height) &&
         (@width == other.width) &&
         (@orientation == other.orientation)
        is_same = true
      end
      is_same
    end

    # @return [True, False] compare dimensions of size, ignoring orientation
    def same_dimensions?(other)
      (@height == other.height) && (@width == other.width) ||
        (@height == other.width) && (@width == other.height)
    end

    # @return [String] get human format name
    def name
      return 'US Letter' if same_dimensions?(Size.new(27.94, 21.59))
      return 'US Legal' if same_dimensions?(Size.new(35.56, 21.59))
      return 'A4' if same_dimensions?(Size.new(29.7, 21.0))
      return 'A5' if same_dimensions?(Size.new(20.99, 14.81))
      return 'B5' if same_dimensions?(Size.new(25.01, 17.6))
      return 'Envelope #10' if same_dimensions?(Size.new(24.13, 10.48))
      return 'Envelope DL' if same_dimensions?(Size.new(22.01, 11.01))
      return 'Tabloid' if same_dimensions?(Size.new(43.17, 27.94))
      return 'A3' if same_dimensions?(Size.new(42.01, 29.7))
      return 'Tabloid Oversize' if same_dimensions?(Size.new(45.71, 30.48))
      return 'ROC 16K' if same_dimensions?(Size.new(27.3, 19.68))
      return 'Envelope Choukei 3' if same_dimensions?(Size.new(23.49, 11.99))
      return 'Super B/A3' if same_dimensions?(Size.new(48.25, 33.02))
      "Unknown page size: Height #{@height} Width #{@width}"
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with Size
    # @return [Size] value of Size
    def self.parse(node)
      size = Size.new
      size.orientation = node.attribute('orient').value.to_sym unless node.attribute('orient').nil?
      # TODO: implement and understand, why 566.929, but not `unit_delimeter`
      size.height = (node.attribute('h').value.to_f / 566.929).round(2)
      size.width = (node.attribute('w').value.to_f / 566.929).round(2)
      size
    end
  end
end
