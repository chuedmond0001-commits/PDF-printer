import time
import heapq
import grpc
import gsi
import google.protobuf.json_format as json_format

def set_value(target_object: str, value_dict: dict, project: gsi.Project, analysis_name: str):
    target_data = gsi.Value()
    target_data.struct_value.update(value_dict)
    project.Set(
        gsi.SetRequest(
            analysis=analysis_name,
            object=target_object,
            data=target_data
        )
    )

def get_value(target_object: str, project: gsi.Project, analysis_name: str) -> dict:
    try:
        gsi_value_object = project.Get(
            gsi.GetRequest(
                analysis=analysis_name,
                object=target_object
            )
        )
        return json_format.MessageToDict(gsi_value_object).get("data", {})
    except Exception:
        return {}

def reveal_object(target_object: str, project: gsi.Project, analysis_name: str):
    val = get_value(target_object, project, analysis_name)
    import pprint
    pprint.pprint(val)

def solve_and_load_results(project: gsi.Project, analysis_name: str, max_retries: int = 3, timeout: int = 300):
    for attempt in range(max_retries):
        try:
            solve_analysis_request = gsi.SolveAnalysesRequest(analyses=[analysis_name])
            load_results_request = gsi.LoadResultsRequest(analysis=analysis_name)
            
            print(f"Solving analysis: {analysis_name} (Attempt {attempt + 1}/{max_retries})")
            project.SolveAnalyses(solve_analysis_request)
            
            print("Solver completed. Loading results...")
            time.sleep(2) 
            
            project.LoadResults(load_results_request)
            print("Results loaded successfully")
            return True
            
        except grpc.RpcError as e:
            print(f"\nError on attempt {attempt + 1}: {e.code()} - {e.details()}")
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 5
                print(f"  Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                print("\nAll retry attempts failed.")
                raise
        except Exception as e:
            print(f"\nUnexpected error on attempt {attempt + 1}: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep((attempt + 1) * 5)
            else:
                raise
    return False

def solve_all_analyses(project: gsi.Project, analysis_names: list[str]):
    print(f"Solving {len(analysis_names)} analyses: {analysis_names}...")
    project.SolveAnalyses(gsi.SolveAnalysesRequest(analyses=analysis_names))
    
    print("Loading results for all analyses...")
    for analysis_name in analysis_names:
        project.LoadResults(gsi.LoadResultsRequest(analysis=analysis_name))

def get_min_fos_across_analyses(project: gsi.Project, analysis_names: list[str], num_critical_surfaces: int = 5) -> tuple[float, int, dict, str, bool]:
    min_fos = float('inf')
    critical_slip, critical_geometry, critical_analysis = None, None, None
    valid_analyses = 0
    error_analyses = []
    
    for analysis_name in analysis_names:
        fos, slip, geometry = get_fos_and_slip(project, analysis_name, num_critical_surfaces)
        
        if fos > 900:
            error_analyses.append(analysis_name)
            continue
        
        valid_analyses += 1
        if fos < min_fos:
            min_fos = fos
            critical_slip = slip
            critical_geometry = geometry
            critical_analysis = analysis_name
            
    all_errors = (valid_analyses == 0)
    return min_fos, critical_slip, critical_geometry, critical_analysis, all_errors

def get_result_values(project: gsi.Project, analysis_name: str, result_type: gsi.ResultType, data_param: gsi.DataParamType, step: int = 1) -> list:
    query_results_request = gsi.QueryResultsRequest(
        analysis=analysis_name, step=step, table=result_type, dataparams=[data_param]
    )
    query_results_response = project.QueryResults(query_results_request)
    return list(query_results_response.results[data_param].values)

def get_fos_and_slip(project: gsi.Project, analysis_name: str, num_critical_surfaces: int = 5) -> tuple[float, int, dict]:  
    query_results_request = gsi.QueryResultsRequest(
        analysis=analysis_name, step=1, table=gsi.ResultType.Slip,
        dataparams=[
            gsi.DataParamType.eRawSlipFOS, gsi.DataParamType.eSlipNum,
            gsi.DataParamType.eSlipCenterX, gsi.DataParamType.eSlipCenterY, gsi.DataParamType.eSlipRadius
        ]
    )
    query_results_response = project.QueryResults(query_results_request)
    fos_list = list(query_results_response.results[query_results_request.dataparams[0]].values)
    slip_list = list(query_results_response.results[query_results_request.dataparams[1]].values)
    
    valid_indices = [i for i, fos in enumerate(fos_list) if fos <= 900]
    
    if not valid_indices:
        critical_indices = heapq.nsmallest(num_critical_surfaces, range(len(fos_list)), key=lambda i: fos_list[i])
        return fos_list[critical_indices[0]], slip_list[critical_indices[0]], {}
        
    critical_indices = heapq.nsmallest(num_critical_surfaces, valid_indices, key=lambda i: fos_list[i])
    idx = critical_indices[0]
    
    critical_slip_geometry = {
        'CenterX': list(query_results_response.results[query_results_request.dataparams[2]].values)[idx],
        'CenterY': list(query_results_response.results[query_results_request.dataparams[3]].values)[idx],
        'Radius': list(query_results_response.results[query_results_request.dataparams[4]].values)[idx]}
        
    return fos_list[idx], slip_list[idx], critical_slip_geometry

def get_fully_specified_slip_fos(project: gsi.Project, analysis_name: str, fully_specified_slip_coords: dict) -> dict:
    for attempt in range(3):
        try:
            query_results_request = gsi.QueryResultsRequest(
                analysis=analysis_name, step=1, table=gsi.ResultType.Slip,
                dataparams=[gsi.DataParamType.eRawSlipFOS, gsi.DataParamType.eSlipNum]
            )
            response = project.QueryResults(query_results_request)
            
            fos_values = list(response.results[gsi.DataParamType.eRawSlipFOS].values)
            slip_numbers = list(response.results[gsi.DataParamType.eSlipNum].values)
            slip_fos_map = {int(num): fos for num, fos in zip(slip_numbers, fos_values)}
            
            return {sid: slip_fos_map.get(sid, 999.0) for sid in fully_specified_slip_coords.keys()}

        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < 2: time.sleep(2)
            else: return {sid: 999.0 for sid in fully_specified_slip_coords.keys()}