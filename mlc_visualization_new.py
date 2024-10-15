import streamlit as st
import plotly.graph_objects as go
import pydicom
import pandas as pd
import os
import subprocess
from convert_xlsx_efs import *


# @st.cache_data
def read_dicom(dicom_content):
    return pydicom.dcmread(dicom_content, force=True)


# @st.cache_data
def plot_beam_eye_view(dicom_content, beam_index, control_point_index):
    ds = read_dicom(dicom_content)

    beam = ds.BeamSequence[beam_index]
    control_point = beam.ControlPointSequence[control_point_index]

    jaws = None
    mlc_positions = None
    try:
        for bl_device in control_point.BeamLimitingDevicePositionSequence:
            if bl_device.RTBeamLimitingDeviceType == 'ASYMX':
                jaws = bl_device.LeafJawPositions
            elif bl_device.RTBeamLimitingDeviceType == 'MLCY':
                mlc_positions = bl_device.LeafJawPositions
    except AttributeError:
        st.error("Selected control point does not contain 'BeamLimitingDevicePositionSequence'.")
        return None

    if jaws is None or mlc_positions is None:
        st.error("Could not find the required ASYMX or MLCY device positions in the control point.")
        return None

    leaf_width_cm = 0.7175
    x_start_cm = -28.7

    fig = go.Figure()

    # Add red axis lines through 0,0
    fig.add_shape(type="line", x0=x_start_cm, y0=0, x1=-x_start_cm, y1=0, line=dict(color="Yellow", width=2))
    fig.add_shape(type="line", x0=0, y0=-200, x1=0, y1=200, line=dict(color="Yellow", width=2))

    fig.add_shape(type="line", x0=jaws[0] / 10, y0=-200, x1=jaws[0] / 10, y1=200, line=dict(color="Red", width=2))
    fig.add_shape(type="line", x0=jaws[1] / 10, y0=-200, x1=jaws[1] / 10, y1=200, line=dict(color="Blue", width=2))

    for i in range(160):  # Assuming 160 leaves total
        x_position_cm = x_start_cm + (i % 80) * leaf_width_cm

        if i < 80:
            fig.add_shape(type="rect",
                          x0=x_position_cm, y0=-200,
                          x1=x_position_cm + leaf_width_cm, y1=mlc_positions[i],
                          line=dict(color="Green"), fillcolor="Green", opacity=0.5)
        else:
            fig.add_shape(type="rect",
                          x0=x_position_cm, y0=200,
                          x1=x_position_cm + leaf_width_cm, y1=mlc_positions[i],
                          line=dict(color="Purple"), fillcolor="Purple", opacity=0.5)

    # Update layout with axis ticks
    fig.update_layout(
        title=f'Beam\'s Eye View: Beam {beam_index + 1}, Control Point {((control_point_index // 2) + 1)}',
        xaxis_title='X Position (cm)',
        yaxis_title='Y Position (mm)',
        xaxis=dict(range=[x_start_cm, x_start_cm + 80 * leaf_width_cm], showgrid=True, dtick=leaf_width_cm * 5),
        yaxis=dict(range=[-200, 200], showgrid=True, dtick=20),
        plot_bgcolor='white'
    )

    return fig


# @st.cache_data
def plot_beam_eye_view_new(dicom_content, beam_index, control_point_index):
    ds = read_dicom(dicom_content)

    beam = ds.BeamSequence[beam_index]
    control_point = beam.ControlPointSequence[control_point_index]

    jaws = None
    mlc_positions = None
    try:
        for bl_device in control_point.BeamLimitingDevicePositionSequence:
            if bl_device.RTBeamLimitingDeviceType == 'ASYMX':
                jaws = bl_device.LeafJawPositions
            elif bl_device.RTBeamLimitingDeviceType == 'MLCY':
                mlc_positions = bl_device.LeafJawPositions
    except AttributeError:
        st.error("Selected control point does not contain 'BeamLimitingDevicePositionSequence'.")
        return None

    if jaws is None or mlc_positions is None:
        st.error("Could not find the required ASYMX or MLCY device positions in the control point.")
        return None

    leaf_width_cm = 0.7175
    x_start_cm = -28.7

    fig = go.Figure()

    # Add axis lines through 0,0
    fig.add_shape(type="line", x0=x_start_cm, y0=0, x1=-x_start_cm, y1=0,
                  line=dict(color="Yellow", width=2))
    fig.add_shape(type="line", x0=0, y0=-200, x1=0, y1=200,
                  line=dict(color="Yellow", width=2))

    # Add jaw positions
    fig.add_shape(type="line", x0=jaws[0] / 10, y0=-200, x1=jaws[0] / 10, y1=200,
                  line=dict(color="Red", width=2))
    fig.add_shape(type="line", x0=jaws[1] / 10, y0=-200, x1=jaws[1] / 10, y1=200,
                  line=dict(color="Blue", width=2))

    # Loop over leaves and add each as an individual Scatter trace
    for i in range(160):  # Assuming 160 leaves total
        x_position_cm = x_start_cm + (i % 80) * leaf_width_cm
        x0 = x_position_cm
        x1 = x_position_cm + leaf_width_cm

        if i < 80:
            y0 = -200
            y1 = mlc_positions[i]
            color = 'rgba(0, 255, 0, 0.5)'  # Green with opacity
            DICOM_leaf_no = i+80
        else:
            y0 = 200
            y1 = mlc_positions[i]
            color = 'rgba(128, 0, 128, 0.5)'  # Purple with opacity
            DICOM_leaf_no = i-80

        x = [x0, x1, x1, x0, x0]
        y = [y0, y0, y1, y1, y0]

        fig.add_trace(go.Scatter(
            x=x,
            y=y,
            fill='toself',
            fillcolor=color,
            line=dict(color='rgba(0,0,0,0.1)'),  # Slightly visible lines
            hoverinfo='text',
            hoveron='fills',
            text=f"Leaf {DICOM_leaf_no}",
            mode='lines',  # Use 'lines' to ensure hover events
            showlegend=False,
        ))

    # Update layout with axis ticks and labels
    fig.update_layout(
        # title=f"Beam's Eye View: Beam {beam_index + 1}, Control Point {((control_point_index // 2) + 1)}",
        title=f"Beam's Eye View: Beam {beam_index + 1}, Control Point {control_point_index}",
        xaxis_title='X Position (cm)',
        yaxis_title='Y Position (mm)',
        xaxis=dict(
            range=[x_start_cm, x_start_cm + 80 * leaf_width_cm],
            showgrid=True,
            dtick=leaf_width_cm * 5
        ),
        yaxis=dict(
            range=[-200, 200],
            showgrid=True,
            dtick=20
        ),
        plot_bgcolor='white'
    )

    return fig


def save_excel(dicom_content, save_path):
    ds = pydicom.dcmread(dicom_content, force=True)
    file_path = os.path.join(save_path, "Beam_Data.xlsx")
    patient_ID = ds.PatientID
    plan_name = ds.RTPlanName
    serial_num = '600074'
    leaf_width = 0.7

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for beam_index, beam in enumerate(ds.BeamSequence):
            data = {}
            beam_name = beam.BeamName
            beam_description = beam.BeamDescription
            for control_point_index, control_point in enumerate(beam.ControlPointSequence):
                try:
                    jaws = None
                    mlc_positions = None
                    gantry_angle = beam.ControlPointSequence[0].GantryAngle  # Extract the gantry angle
                    gantry_rotation = [beam.ControlPointSequence[0].GantryRotationDirection]
                    cumulative_meterset = [round(float(control_point.CumulativeMetersetWeight * 100), 3)]
                    beam_MU = [ds.FractionGroupSequence[0].ReferencedBeamSequence[beam_index].BeamMeterset]

                    for bl_device in control_point.BeamLimitingDevicePositionSequence:
                        if bl_device.RTBeamLimitingDeviceType == 'ASYMX':
                            jaws = [x / 10 for x in bl_device.LeafJawPositions][::-1]
                            jaws = [round(elem, 2) for elem in jaws]
                        elif bl_device.RTBeamLimitingDeviceType == 'MLCY':
                            mlc_positions = [x / 10 for x in bl_device.LeafJawPositions][::-1]
                            # Reverse and divide by 10
                            Y1 = [round(elem, 3) for elem in mlc_positions[80:160]]
                            Y2 = [round(elem, 3) for elem in mlc_positions[0:80]]

                    if jaws and mlc_positions:
                        data[f'Control Point {control_point_index + 1}'] = beam_MU + [serial_num] + [patient_ID] + \
                                                                           [plan_name] + [beam_name] \
                                                                           + [beam_description] + \
                                                                           [leaf_width] + \
                                                                           cumulative_meterset + [gantry_angle] + \
                                                                           jaws + [""] + \
                                                                           gantry_rotation + Y1 + Y2
                except AttributeError:
                    print(gantry_angle)
                    continue

            if data:
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=f'Beam {beam_index + 1}', index=False)

    st.success(f"Excel file saved at {file_path}")
    return file_path


def main():
    st.title("Unity MLC Explorer + RTP Mangler")
    st.write(
        "Open RT Plan DICOM and view MLC sequences as Excel or .efs files for use with iComCat."
        "  \n"
        "Modify with RTP Mangler")
    uploaded_file = st.file_uploader("Choose a DICOM file", type=["dcm"])
    if uploaded_file is not None:
        save_directory = st.text_input("Enter the directory to save the Excel file:", value=os.getcwd())

        if 'excel_file_path' not in st.session_state:
            st.session_state.excel_file_path = None

        if st.button("Save All Beams to Excel"):
            st.session_state.excel_file_path = save_excel(uploaded_file, save_directory)

        if st.session_state.excel_file_path:
            st.write("## Excel File Content")
            df = pd.read_excel(st.session_state.excel_file_path, sheet_name=None)
            beam_options = list(df.keys())
            selected_beam = st.selectbox("Select Beam", options=beam_options, key="beam_select")
            st.write(df[selected_beam])

        if st.button("Generate .efs File"):
            if st.session_state.excel_file_path:
                efs_file_path = os.path.join(save_directory)
                process_excel_to_text_sheets(st.session_state.excel_file_path, efs_file_path)
                st.success(f".efs file generated at {efs_file_path}")
            else:
                st.error("Please save the Excel file first.")

        ds = pydicom.dcmread(uploaded_file, force=True)
        beam_options = range(len(ds.BeamSequence))
        selected_beam_index = st.selectbox("Select Beam", options=beam_options, format_func=lambda x: f"Beam {x + 1}",
                                           key="dicom_beam_select")

        control_points = ds.BeamSequence[selected_beam_index].ControlPointSequence
        valid_control_points = [
            i for i, cp in enumerate(control_points)
            if i % 2 != 0 and hasattr(cp, 'BeamLimitingDevicePositionSequence')
        ]

        if len(valid_control_points) == 1:
            control_point_index = valid_control_points[0]
            st.write(f"Only one valid control point found: Control Point {((control_point_index // 2) + 1)}")
            fig = plot_beam_eye_view_new(uploaded_file, selected_beam_index, control_point_index)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        elif valid_control_points:
            control_point_index = st.select_slider("Select Control Point",
                                                   options=[(i // 2) + 1 for i in valid_control_points],
                                                   key="control_point_select")
            fig = plot_beam_eye_view_new(uploaded_file, selected_beam_index,
                                         valid_control_points[control_point_index - 1])
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.error("No valid odd-indexed control points with 'BeamLimitingDevicePositionSequence' found.")

        st.write("## Modify DICOM File")
        command_string = st.text_area("Enter Command String", value="mu=100 b0 cp0")

        if st.button("Modify DICOM"):
            modified_dicom_path = os.path.join(save_directory, "modified_output.dcm")
            script_path = os.path.join(os.path.dirname(__file__), 'mangle.py')  # Adjust this if mangle.py is in a
            # different directory

            # Save the uploaded file to a temporary location
            temp_dicom_path = os.path.join(save_directory, uploaded_file.name)
            with open(temp_dicom_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Debugging: Check if the script path exists
            if not os.path.exists(script_path):
                st.error(f"Script not found: {script_path}")
            else:
                st.write(f"Script found: {script_path}")

            # Debugging: Check if the uploaded file is saved correctly
            if not os.path.exists(temp_dicom_path):
                st.error(f"Uploaded file not found: {temp_dicom_path}")
            else:
                st.write(f"Uploaded file found: {temp_dicom_path}")

            args = [
                'python', script_path,
                temp_dicom_path,
                '-o', modified_dicom_path,
                command_string
            ]

            if st.checkbox("Verbose Output"):
                args.append('-v')

            if st.checkbox("Keep Original UID"):
                args.append('-k')

            result = subprocess.run(args, capture_output=True, text=True)
            if result.returncode == 0:
                st.success(f"Modified DICOM file saved at {modified_dicom_path}")
            else:
                st.error(f"Error modifying DICOM file: {result.stderr}")


if __name__ == "__main__":
    main()
