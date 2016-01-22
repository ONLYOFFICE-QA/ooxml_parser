module OoxmlParser
  class Configuration
    attr_accessor :units

    def initialize
      @units = :centimeters
    end

    # @return [Float] by which divide units values
    def units_delimiter
      return (20 * 635.0) if @units == :points
      return (566.929 * 635.0) if @units == :centimeters
      return 1 if @units == :dxa
      return 1440 if @units == :inches
      return (1.0 / 635.0) if @units == :emu
      warn "Cannot recognize #{@units} unit. Will use dxa by default"
      1
    end
  end
end
