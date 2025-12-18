import * as React from "react";
import "./Datepicker.css";
import { DateRange } from "./types";

interface DatePickerProps {
  dateRange: DateRange | undefined;
  onDateRangeSelect: (range: DateRange) => void;
  onClose: () => void;
}

const DatePicker: React.FC<DatePickerProps> = ({
  dateRange,
  onDateRangeSelect,
  onClose,
}) => {
  const [currentMonth, setCurrentMonth] = React.useState<Date>(() => {
    return dateRange?.startDate || new Date();
  });

  const [viewMode, setViewMode] = React.useState<"month" | "year">("month");
  const [tempStartDate, setTempStartDate] = React.useState<Date | null>(
    dateRange?.startDate || null
  );
  const [tempEndDate, setTempEndDate] = React.useState<Date | null>(
    dateRange?.endDate || null
  );

  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const monthNamesShort = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  const getDaysInMonth = (date: Date): Date[] => {
    const year = date.getFullYear();
    const month = date.getMonth();
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();

    const days: Date[] = [];

    const firstDayOfWeek = firstDay.getDay();
    for (let i = 0; i < firstDayOfWeek; i++) {
      days.push(new Date(0)); // Empty cell marker
    }

    for (let day = 1; day <= daysInMonth; day++) {
      days.push(new Date(year, month, day, 12, 0, 0, 0));
    }

    return days;
  };

  const days = getDaysInMonth(currentMonth);
  const currentYear = currentMonth.getFullYear();
  const currentMonthIndex = currentMonth.getMonth();

  const goToPreviousMonth = () => {
    setCurrentMonth(
      new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1)
    );
  };

  const goToNextMonth = () => {
    setCurrentMonth(
      new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1)
    );
  };

  const goToPreviousYear = () => {
    setCurrentMonth(
      new Date(currentMonth.getFullYear() - 1, currentMonth.getMonth(), 1)
    );
  };

  const goToNextYear = () => {
    setCurrentMonth(
      new Date(currentMonth.getFullYear() + 1, currentMonth.getMonth(), 1)
    );
  };

  const goToToday = () => {
    const today = new Date();
    setCurrentMonth(new Date(today.getFullYear(), today.getMonth(), 1));
    setViewMode("month");
  };

  const handleDateClick = (date: Date) => {
    if (date.getTime() === 0) return;

    // If no start date, set it
    if (!tempStartDate) {
      setTempStartDate(date);
      return;
    }

    // If start date is set but no end date
    if (!tempEndDate) {
      // If clicked date is before start date, swap them
      if (date < tempStartDate) {
        setTempEndDate(tempStartDate);
        setTempStartDate(date);
      } else {
        setTempEndDate(date);
      }
      return;
    }

    // If both dates are set, reset and start new selection
    setTempStartDate(date);
    setTempEndDate(null);
  };

  const handleApplyRange = () => {
    if (tempStartDate && tempEndDate) {
      onDateRangeSelect({
        startDate: tempStartDate,
        endDate: tempEndDate,
      });
      onClose();
    }
  };

  const handleClearRange = () => {
    setTempStartDate(null);
    setTempEndDate(null);
  };

  const handleMonthClick = (monthIndex: number) => {
    setCurrentMonth(new Date(currentYear, monthIndex, 1));
    setViewMode("month");
  };

  const isToday = (date: Date): boolean => {
    if (date.getTime() === 0) return false;
    const today = new Date();
    return (
      date.getDate() === today.getDate() &&
      date.getMonth() === today.getMonth() &&
      date.getFullYear() === today.getFullYear()
    );
  };

  const isStartDate = (date: Date): boolean => {
    if (date.getTime() === 0 || !tempStartDate) return false;
    return (
      date.getDate() === tempStartDate.getDate() &&
      date.getMonth() === tempStartDate.getMonth() &&
      date.getFullYear() === tempStartDate.getFullYear()
    );
  };

  const isEndDate = (date: Date): boolean => {
    if (date.getTime() === 0 || !tempEndDate) return false;
    return (
      date.getDate() === tempEndDate.getDate() &&
      date.getMonth() === tempEndDate.getMonth() &&
      date.getFullYear() === tempEndDate.getFullYear()
    );
  };

  const isInRange = (date: Date): boolean => {
    if (date.getTime() === 0 || !tempStartDate) return false;
    if (!tempEndDate) return false;

    return date >= tempStartDate && date <= tempEndDate;
  };

  const isCurrentMonth = (monthIndex: number): boolean => {
    const today = new Date();
    return (
      monthIndex === today.getMonth() &&
      currentYear === today.getFullYear()
    );
  };

  const isSelectedMonth = (monthIndex: number): boolean => {
    return monthIndex === currentMonthIndex;
  };

  const formatDateDisplay = (date: Date | null): string => {
    if (!date) return "";
    return date.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  };

  return (
    <div className="date-picker-overlay" onClick={onClose}>
      <div className="date-picker-container" onClick={(e) => e.stopPropagation()}>
        
        {/* Left Panel - Month View */}
        <div className="date-picker-left">
          <div className="date-picker-header">
            <div className="date-picker-month-title">
              {monthNames[currentMonthIndex]} {currentYear}
            </div>
            <div className="date-picker-nav-buttons">
              <button 
                className="date-picker-nav-btn" 
                onClick={goToPreviousMonth}
                title="Previous month"
              >
                ↑
              </button>
              <button 
                className="date-picker-nav-btn" 
                onClick={goToNextMonth}
                title="Next month"
              >
                ↓
              </button>
            </div>
          </div>

          <div className="date-picker-weekdays">
            {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((day) => (
              <div key={day} className="date-picker-weekday">
                {day}
              </div>
            ))}
          </div>

          <div className="date-picker-days">
            {days.map((date, index) => {
              const isEmpty = date.getTime() === 0;
              const isOutsideMonth = !isEmpty && date.getMonth() !== currentMonthIndex;
              const startDte = isStartDate(date);
              const endDte = isEndDate(date);
              const inRange = isInRange(date);
              
              return (
                <button
                  key={index}
                  className={`date-picker-day ${
                    isEmpty ? "date-picker-day-empty" : ""
                  } ${isOutsideMonth ? "date-picker-day-outside" : ""} ${
                    isToday(date) ? "date-picker-day-today" : ""
                  } ${startDte ? "date-picker-day-start" : ""} ${
                    endDte ? "date-picker-day-end" : ""
                  } ${inRange && !startDte && !endDte ? "date-picker-day-in-range" : ""}`}
                  onClick={() => handleDateClick(date)}
                  disabled={isEmpty}
                >
                  {isEmpty ? "" : date.getDate()}
                </button>
              );
            })}
          </div>

          <div className="date-picker-footer">
            <button className="date-picker-today-btn" onClick={goToToday}>
              Today
            </button>
          </div>
        </div>

        {/* Right Panel - Summary & Controls */}
        <div className="date-picker-right">
          <div className="date-picker-year-header">
            <div className="date-picker-year-title">
              Custom Date Range
            </div>
          </div>

          <div className="date-picker-range-summary">
            <div className="date-picker-range-item">
              <div className="date-picker-range-label">From:</div>
              <div className="date-picker-range-value">
                {formatDateDisplay(tempStartDate) || "Not selected"}
              </div>
            </div>
            <div className="date-picker-range-item">
              <div className="date-picker-range-label">To:</div>
              <div className="date-picker-range-value">
                {formatDateDisplay(tempEndDate) || "Not selected"}
              </div>
            </div>
          </div>

          {tempStartDate && tempEndDate && (
            <div className="date-picker-range-info">
              <span>
                {Math.floor((tempEndDate.getTime() - tempStartDate.getTime()) / (1000 * 60 * 60 * 24)) + 1} days selected
              </span>
            </div>
          )}

          <div className="date-picker-button-group">
            <button
              className="date-picker-apply-btn"
              onClick={handleApplyRange}
              disabled={!tempStartDate || !tempEndDate}
            >
              Apply Range
            </button>
            <button
              className="date-picker-clear-btn"
              onClick={handleClearRange}
              disabled={!tempStartDate && !tempEndDate}
            >
              Clear
            </button>
          </div>

          <button 
            className="date-picker-close-btn" 
            onClick={onClose}
          >
            ✕
          </button>
        </div>
      </div>
    </div>
  );
};

export default DatePicker;