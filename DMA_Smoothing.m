%% User Input
% The file should be an Excel file, exported from TRIOS software as an .xls format.
file_name = input('Enter the Excel file name (with extension): ', 's');
sheet_name = input('Enter the sheet name: ', 's');

%% Read Data
% This section reads data from the specified Excel file and sheet.
% It assumes that column 3 contains temperature values and column 6 contains tan(δ) values.
% Adjust these column indices if the data is organized differently.
data = readtable(file_name, 'Sheet', sheet_name);

% Check if the required columns are available in the data.
% If not, display an error message and stop execution.
if size(data, 2) < 6
    error('The specified file does not have enough columns for processing.');
end

%% Extract Data Columns
% Extract temperature values from column 3 and tan(δ) values from column 6 of the data.
temperature = data{:, 3};
tan_delta = data{:, 6};

%% Smoothing
% Define the window size for moving average smoothing.
% The window size (moving_avg_window) determines the amount of data averaging.
moving_avg_window = 15;

% Apply the custom smoothing function 'process' to the temperature and tan(δ) data.
% The function performs data cleaning, interpolation, and smoothing.
[temperature_smoothed, tan_delta_smoothed] = process(temperature, tan_delta, moving_avg_window);

%% Plotting
figure;

% Plot the original data (unsmoothed) in light grey.
plot(temperature, tan_delta, 'Color', [0.7 0.7 0.7], 'LineWidth', 1.5, 'DisplayName', 'Original Data');
hold on;

% Plot the smoothed data in a contrasting color (blue).
plot(temperature_smoothed, tan_delta_smoothed, 'Color', [0.2 0.4 0.8], 'LineWidth', 2.5, 'DisplayName', 'Smoothed Data');

% Add labels for the x and y axes.
xlabel('Temperature (^{o}C)', 'FontName', 'Arial', 'FontSize', 15);
ylabel('tan(δ)', 'FontName', 'Arial', 'FontSize', 15);

% Display a legend in the Northwest position to differentiate between the original and smoothed data.
legend('show', 'Location', 'Northwest', 'FontName', 'Arial', 'FontSize', 12);

% Adjust the x and y axis limits to encompass all data points.
xlim([min(temperature) max(temperature)]);
ylim([min(tan_delta) max(tan_delta)]);

% Format the plot appearance for better readability.
box off;
ax = gca;
ax.XAxis.LineWidth = 1.2;
ax.YAxis.LineWidth = 1.2;
set(gca, 'TickDir', 'out');
legend boxoff;

%% Smoothing Function
% This function takes in original x and y data, removes any invalid entries,
% averages duplicate x-values, interpolates to create an evenly spaced x-axis,
% and applies a moving average smoothing to reduce noise in the y-data.
% The function outputs smoothed x and y values for plotting.
function [x_new, y_smoothed] = process(x_original, y_original, moving_avg_window)
    % Remove invalid data points (e.g., NaN or Inf values) to prevent errors in processing.
    valid_indices = isfinite(x_original) & isfinite(y_original);
    x_original = x_original(valid_indices);
    y_original = y_original(valid_indices);

    % Identify unique x-values and average the corresponding y-values for duplicate x-values.
    [x_unique, ~, idx] = unique(x_original);
    y_unique = accumarray(idx, y_original, [], @mean);

    % Interpolate the unique x and y values to create a uniformly spaced x-axis.
    % This step is necessary for smoothing and creates a new x-axis spanning the range of the original data.
    x_new = linspace(min(x_unique), max(x_unique), length(x_unique));
    y_interpolated = interp1(x_unique, y_unique, x_new, 'linear');

    % Apply a moving average to the interpolated y values to smooth out noise.
    % The size of the moving average window is defined by 'moving_avg_window'.
    y_smoothed = movmean(y_interpolated, moving_avg_window);
end
